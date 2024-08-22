using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace Excel_to_JSON_converter
{

    internal class DocAgentInfo
    {
        [JsonExtensionData]
        public Dictionary<string, object> info { get; set; }
        public DocAgentInfo()
        {
            info = new();
            info.Add("vfam", null);
            info.Add("vname", null);
            info.Add("votch",null);
            info.Add("cvdoc", "01");
            info.Add("cln", null);
            info.Add("cstranf", "112");
            info.Add("nrate", 13);
        }
    }
}
