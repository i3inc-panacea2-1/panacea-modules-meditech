using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace UserPlugins.Meditech.Models
{
	[DataContract]
	public class GetMeditechSettingsResponse
	{
		[DataMember(Name = "timeout")]
		public int Timeout { get; set; }

        [DataMember(Name = "paths")]
        public List<string> Paths { get; set; }
    }

	
}
