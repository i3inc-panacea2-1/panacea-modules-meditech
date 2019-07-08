using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Threading;
using Microsoft.Win32;
using Application = System.Windows.Application;
using UserPlugins.Meditech.Models;
using System.IO;
using System.Windows.Automation;
using System.Management;
using Windows.UserActivity;
using Panacea.Modularity.UiManager;
using Panacea.Core;
using System.Web;
using Panacea.Modularity.AppBar;

namespace Panacea.Modules.Meditech
{
    public class MeditechPage : ICallablePlugin
    {
        private readonly PanaceaServices _core;
        private readonly ILogger _logger;
        readonly LoadingWindow _loadingwindow;

        DispatcherTimer _checker;
        GetMeditechSettingsResponse _settings;

        public MeditechPage(PanaceaServices core)
        {
            _core = core;
            _logger = core.Logger;
            _loadingwindow = new LoadingWindow();
            _timer = new IdleTimer(TimeSpan.FromSeconds(Timeout));
            _timer.Tick += _timer_Elapsed;

            _checker = new DispatcherTimer();
            _checker.Interval = TimeSpan.FromSeconds(3);
            _checker.Tick += _checker_Tick;
        }

        private void _checker_Tick(object sender, EventArgs e)
        {
            if (!IsMeditechRunning())
            {
                _timer.Stop();
                _checker.Stop();
                if (_core.TryGetUiManager(out IUiManager ui))
                {
                    Application.Current.Dispatcher.Invoke(ui.Resume);
                }
                return;
            }
            if (MeditechHasUi())
            {
                Application.Current.Dispatcher.Invoke(() =>
                {
                    _loadingwindow.Hide();
                });
            }
            else
            {
                Application.Current.Dispatcher.Invoke(() =>
                {
                    _loadingwindow.Show();
                });

            }
            MaximizeMeditech();
        }

        private int Timeout = 600;

        void _timer_Elapsed(object sender, EventArgs e)
        {
            if (IsMeditechRunning())
            {
                _timer.Stop();
                _checker.Stop();
                CloseMeditech();
                if (_core.TryGetUiManager(out IUiManager ui))
                {
                    Application.Current.Dispatcher.Invoke(ui.Resume);
                }
                return;
            }
        }

        private const int SW_SHOWNORMAL = 1;
        private const int SW_SHOWMINIMIZED = 2;
        private const int SW_SHOWMAXIMIZED = 3;

        [DllImport("user32.dll")]
        private static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);

        void MaximizeMeditech()
        {
            var procs = Process.GetProcessesByName("Magic");
            foreach (var proc in procs)
            {
                try
                {
                    if (proc.MainWindowHandle != IntPtr.Zero)
                    {
                        ShowWindowAsync(proc.MainWindowHandle, SW_SHOWMAXIMIZED);
                    }
                }
                catch
                {
                }
            }
        }

        bool IsMeditechRunning()
        {
            var procs = Process.GetProcessesByName("MTAppDwn");
            if (procs.Length == 0) return false;
            var procs1 = Process.GetProcessesByName("Magic");
            var hasWindows = false;
            foreach (var pr in procs)
                if (pr.MainWindowHandle != IntPtr.Zero) hasWindows = true;

            return hasWindows || procs1.Length > 0;

        }

        bool MeditechHasUi()
        {
            var procs = Process.GetProcessesByName("MTAppDwn");
            var procs1 = Process.GetProcessesByName("Magic");
            return procs.Any(p => IsProcessWindowed(p)) || procs1.Any(p => IsProcessWindowed(p));

        }

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool GetWindowPlacement(IntPtr hWnd, ref WINDOWPLACEMENT lpwndpl);

        private struct WINDOWPLACEMENT
        {
            public int length;
            public int flags;
            public int showCmd;
            public System.Drawing.Point ptMinPosition;
            public System.Drawing.Point ptMaxPosition;
            public System.Drawing.Rectangle rcNormalPosition;
        }




        private Boolean IsProcessWindowed(Process externalProcess)
        {

            try
            {
                foreach (ProcessThread threadInfo in externalProcess.Threads)
                {
                    IntPtr[] windows = GetWindowHandlesForThread(threadInfo.Id);

                    if (windows != null)
                    {
                        foreach (IntPtr handle in windows)
                        {
                            if (IsWindowVisible(handle))
                            {
                                WINDOWPLACEMENT placement = new WINDOWPLACEMENT();
                                GetWindowPlacement(handle, ref placement);
                                _logger.Debug(this, $"{placement.showCmd} {placement.rcNormalPosition.Height} {placement.rcNormalPosition.Size.Height}");
                                if (placement.showCmd == 2) continue;
                                if (placement.rcNormalPosition.Height < 80) continue;
                                return true;
                            }
                        }
                    }
                }
            }
            catch { }
            return false;
        }

        private IntPtr[] GetWindowHandlesForThread(int threadHandle)
        {
            results.Clear();
            EnumWindows(WindowEnum, threadHandle);

            return results.ToArray();
        }

        private delegate int EnumWindowsProc(IntPtr hwnd, int lParam);

        private List<IntPtr> results = new List<IntPtr>();

        private int WindowEnum(IntPtr hWnd, int lParam)
        {
            int processID = 0;
            int threadID = GetWindowThreadProcessId(hWnd, out processID);
            if (threadID == lParam)
            {
                results.Add(hWnd);
            }

            return 1;
        }

        [DllImport("user32.Dll")]
        private static extern int EnumWindows(EnumWindowsProc x, int y);
        [DllImport("user32.dll")]
        public static extern int GetWindowThreadProcessId(IntPtr handle, out int processId);
        [DllImport("user32.dll")]
        static extern bool IsWindowVisible(IntPtr hWnd);
        void CloseMeditech()
        {
            _loadingwindow.Hide();

            var procs = Process.GetProcessesByName("MTAppDwn");
            foreach (var proc in procs)
            {
                if (IsService(proc.Id)) continue;
                try
                {
                    proc.Kill();
                }
                catch
                {
                }
            }

            procs = Process.GetProcessesByName("Magic");
            foreach (var proc in procs)
            {
                try
                {
                    proc.Kill();
                }
                catch
                {
                }
            }

            procs = Process.GetProcessesByName("VMagicPPVW");
            foreach (var proc in procs)
            {
                try
                {
                    proc.Kill();
                }
                catch
                {
                }
            }
        }

        public static bool IsService(int pID)
        {
            using (ManagementObjectSearcher Searcher = new ManagementObjectSearcher(
            "SELECT * FROM Win32_Service WHERE ProcessId =" + "\"" + pID + "\""))
            {
                foreach (ManagementObject service in Searcher.Get())
                    return true;
            }
            return false;
        }

        private IdleTimer _timer;


        bool _accepting = false;

        public async Task BeginInit()
        {

            //todo _comm.RegisterUri("meditech", OnUri);

            try
            {
                var response = await _core.HttpClient.GetObjectAsync<GetMeditechSettingsResponse>("meditech/get_settings/");
                if (response.Success)
                {
                    _settings = response.Result;
                    Timeout = response.Result.Timeout;
                    _timer?.Stop();
                    _timer = new IdleTimer(TimeSpan.FromSeconds(Timeout));
                    _timer.Tick += _timer_Elapsed;
                }
            }
            catch
            {

            }

        }

        private object OnUri(Uri uri)
        {
            Task.Run(async () =>
            {
                await OnUriAsync(uri);
            });

            return null;
        }

        private async Task OnUriAsync(Uri uri)
        {
            if (_accepting) return;
            _accepting = true;
            var task = Task.Run(async () =>
            {
                await Task.Delay(2000);
                _accepting = false;
            });
            //todo _webSocket.PopularNotifyPage("Meditech");

            var args = HttpUtility.ParseQueryString(uri.Query);
            var username = args["username"];
            var password = args["password"];
            var domain = args["domain"];
            if (IsMeditechRunning())
            {
                CloseMeditech();
                return;
            }
            _logger.Debug(this, $"Opening with username: {username}, password (len): {password.Length}, domain: {domain}");

            try
            {


                Application.Current.Dispatcher.Invoke(() =>
                {
                    OpenMeditech(true);
                });
                await Task.Delay(400);
                var w = FindWindow(null, "Enter Network Password");
                var i = 0;
                while (w == IntPtr.Zero && i < 1000)
                {
                    await Task.Delay(300);
                    w = FindWindow(null, "Enter Network Password");
                    i++;
                }
                if (w == IntPtr.Zero)
                {
                    _logger.Debug(this, "Meditech window not found");
                    CloseMeditech();
                    return;
                }
                var window = AutomationElement.FromHandle(w);
                //MessageBox.Show(window.Current.Name);
                CacheRequest cacheRequest = new CacheRequest();
                cacheRequest.Add(AutomationElement.NameProperty);
                cacheRequest.Add(AutomationElement.HasKeyboardFocusProperty);
                cacheRequest.TreeScope = TreeScope.Element | TreeScope.Children;


                var list = window.GetUpdatedCache(cacheRequest);

                var usernameBox = list.FindFirst(TreeScope.Children, new AndCondition(
                   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit),
                   new PropertyCondition(AutomationElement.IsPasswordProperty, false)));

                _logger.Debug(this, "Setting username box focus");
                object valuePattern = null;

                if (!usernameBox.TryGetCurrentPattern(
                    ValuePattern.Pattern, out valuePattern))
                {
                    _logger.Debug(this, "No ValuePattern for username box");
                }
                else
                {
                    if (((ValuePattern)valuePattern).Current.IsReadOnly)
                    {
                        throw new InvalidOperationException(
                            "The control is read-only.");
                    }
                    else
                    {
                        _logger.Debug(this, "Typing username");
                        ((ValuePattern)valuePattern).SetValue(username);
                    }
                }


                var passwordBox = list.FindFirst(TreeScope.Children, new AndCondition(
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit),
                    new PropertyCondition(AutomationElement.IsPasswordProperty, true)));

                _logger.Debug(this, "Setting password box focus");
                valuePattern = null;

                if (!passwordBox.TryGetCurrentPattern(
                    ValuePattern.Pattern, out valuePattern))
                {
                    _logger.Debug(this, "No ValuePattern for password box");
                }
                else
                {
                    if (((ValuePattern)valuePattern).Current.IsReadOnly)
                    {
                        throw new InvalidOperationException(
                            "The control is read-only.");
                    }
                    else
                    {
                        _logger.Debug(this, "Typing password");
                        ((ValuePattern)valuePattern).SetValue(password);
                    }
                }
                var domainBox = list.FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ComboBox));

                var max = 0.0;
                string value = null;
                var items = GetComboboxItems(domainBox);

                foreach (var option in items)
                {
                    var current = domain.LongestCommonSubsequence(option).Item2;
                    if (current > max)
                    {
                        max = current;
                        value = option;
                    }
                }
                _logger.Debug(this, "Setting domain " + value);
                if (value != null)
                    SelectComboboxItem(domainBox, value);

                if (args["AutoLogin"] != null && args["autoLogin"].ToLower() == "true")
                {
                    var okButton = list.FindFirst(TreeScope.Children, new AndCondition(
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button),
                    new PropertyCondition(AutomationElement.NameProperty, "OK")));
                    await Task.Delay(360);
                    var invokePattern = okButton.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
                    _logger.Debug(this, "Pressing OK");
                    invokePattern.Invoke();
                }


            }
            catch (Exception ex)
            {
                _logger.Error(this, ex.Message);
            }
        }

        private static int LevenshteinDistance(string s, string t)
        {
            int n = s.Length;
            int m = t.Length;
            int[,] d = new int[n + 1, m + 1];
            if (n == 0)
            {
                return m;
            }
            if (m == 0)
            {
                return n;
            }
            for (int i = 0; i <= n; d[i, 0] = i++)
                ;
            for (int j = 0; j <= m; d[0, j] = j++)
                ;
            for (int i = 1; i <= n; i++)
            {
                for (int j = 1; j <= m; j++)
                {
                    int cost = (t[j - 1] == s[i - 1]) ? 0 : 1;
                    d[i, j] = Math.Min(
                        Math.Min(d[i - 1, j] + 1, d[i, j - 1] + 1),
                        d[i - 1, j - 1] + cost);
                }
            }
            return d[n, m];
        }

        private static IEnumerable<string> GetComboboxItems(AutomationElement comboBox)
        {
            if (comboBox == null) yield break;
            // Get the list box within the combobox
            AutomationElement listBox = comboBox.FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.List));
            if (listBox == null) yield break;
            // Search for item within the listbox
            var listItem = listBox.FindAll(TreeScope.Children, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ListItem));
            if (listItem == null) yield break;
            foreach (AutomationElement item in listItem)
            {
                CacheRequest req = new CacheRequest();
                req.Add(AutomationElement.NameProperty);
                yield return item.GetUpdatedCache(req).Cached.Name;
            }
        }

        private static bool SelectComboboxItem(AutomationElement comboBox, string item)
        {
            if (comboBox == null) return false;
            // Get the list box within the combobox
            AutomationElement listBox = comboBox.FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.List));
            if (listBox == null) return false;
            // Search for item within the listbox
            AutomationElement listItem = listBox.FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.NameProperty, item));
            if (listItem == null) return false;
            // Check if listbox item has SelectionItemPattern
            object objPattern;
            if (true == listItem.TryGetCurrentPattern(SelectionItemPatternIdentifiers.Pattern, out objPattern))
            {
                SelectionItemPattern selectionItemPattern = objPattern as SelectionItemPattern;
                selectionItemPattern.Select(); // Invoke Selection
                return true;
            }
            return false;
        }

        [DllImport("user32")]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        void Host_Resumed(object sender, EventArgs e)
        {
            CloseMeditech();
        }

        public Task EndInit()
        {
            if (_core.TryGetUiManager(out IUiManager ui))
            {
                ui.Resumed += Host_Resumed;
            }
            return Task.CompletedTask;
        }

        public Task Shutdown()
        {
            return Task.CompletedTask;
        }

        public async void Call()
        {
            try
            {
                if(_core.TryGetAppBar(out IAppBar bar))
                {
                    bar.Show();
                    return;
                }
                //_webSocket.PopularNotifyPage("Meditech");
                OpenMeditech(true);
                await Task.Delay(1000);
                _timer.Start();
                _checker.Start();
            }
            catch
            {
                //Host.Window.SwitchToAppBar(null);
            }
        }

        private async void OpenMeditech(bool deleteLastUser = false)
        {
            if (_core.TryGetUiManager(out IUiManager ui))
            {
                if (deleteLastUser)
                {
                    using (var key = Registry.CurrentUser.OpenSubKey(@"Software\Meditech\MTAppDwn", true))
                    {
                        if (key != null)
                            key.SetValue("LastUser", "");
                    }
                }
                DeleteMeditechData();
                var path = GetMeditechPath();
                if (path == null || !File.Exists(path))
                {
                    ui.Toast("Meditech not found");
                    return;
                }
                Process.Start(path);
                ui.Pause();
                var w = FindWindow(null, "Enter Network Password");
                var i = 0;
                while (w == IntPtr.Zero && i < 1000)
                {
                    await Task.Delay(300);
                    w = FindWindow(null, "Enter Network Password");
                    i++;
                }
                if (w == IntPtr.Zero)
                {
                    //Host.Logger.Debug(this, "Meditech window not found");
                    CloseMeditech();
                    return;
                }
                _timer.Start();
                _checker.Start();
            }
        }

        private void DeleteMeditechData()
        {
            //return;
            //try
            //{
            //    var root = Path.Combine(@"C:\ProgramData\MEDITECH");
            //    var universe = Directory.GetDirectories(root).First(p => p.EndsWith(".Universe"));
            //    var ring = Directory.GetDirectories(universe).First(p => p.EndsWith(".LIVEF.Ring"));
            //    Directory.Delete(ring, true);
            //}
            //catch
            //{

            //}
        }

        private string GetMeditechPath()
        {
            return _settings?.Paths?.FirstOrDefault(p => File.Exists(p));
        }

        public void Dispose()
        {

        }
    }
}
