using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace Excel_to_JSON_converter
{
    internal class PckAgentInfo
    {
        public DateTime dcreate { get; set; }

        [JsonExtensionData]
        public Dictionary<string, object> AdditionalFields { get; set; }
        public PckAgentInfo(DateTime date)
        {
            dcreate = date;
            AdditionalFields = new();
            AdditionalFields.Add("ndepno", (int)0);
            AdditionalFields.Add("ngod", (int)0);
            AdditionalFields.Add("ntype", (int)0);
            AdditionalFields.Add("vexec", "");
            AdditionalFields.Add("vunp", "");
            AdditionalFields.Add("vphn", "");
            AdditionalFields.Add("nmns", (int)0);
            AdditionalFields.Add("nmnsf", (int)0);
        }

    }
}
