
using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace Excel_to_JSON_converter
{
    internal class DocAgent
    {
        public DocAgentInfo docagentinfo { get; set; }
        public LinkedList<Tar4>? tar4 { get; set; }
        public LinkedList<Tar5>? tar5 { get; set; }
        public LinkedList<Tar7>? tar7 { get; set; }
        public LinkedList<Tar14>? tar14 { get; set; }

        [JsonExtensionData]
        public Dictionary<string,object> AdditionalFields { get; set; }
        public DocAgent(DocAgentInfo agentInfo)
        {
            docagentinfo = agentInfo;
            AdditionalFields = new();
            AdditionalFields.Add("ntsumincome", (double)0);
            AdditionalFields.Add("ntsumexemp", (double)0);
            AdditionalFields.Add("ntsumnotcalc", (double)0);
            AdditionalFields.Add("nsumstand", (double)0);
            AdditionalFields.Add("ntsumsoc", (double)0);
            AdditionalFields.Add("ntsumprop", (double)0);
            AdditionalFields.Add("ntsumprof", (double)0);
            AdditionalFields.Add("ntsumsec", (double)0);
            AdditionalFields.Add("ntsumtrust", (double)0);
            AdditionalFields.Add("ntsumbank", (double)0);
            AdditionalFields.Add("ntsumcalcincome", (double)0);
            AdditionalFields.Add("ntsumcalcincomediv", (double)0);
            AdditionalFields.Add("ntsumwithincome", (double)0);
            AdditionalFields.Add("ntsumwithincomediv", (double)0);
        }
    }
}
