using System.Text.Json.Serialization;
namespace Excel_to_JSON_converter
{
    internal class PckAgent
    {
        [JsonIgnore]
        public Dictionary<string, DocAgent> docagentMap { get; set; }
        public PckAgentInfo pckagentinfo { get; set; }
        public List<DocAgent> docagent
        {
            get { return docagentMap.Values.ToList(); }
        }
        public PckAgent(PckAgentInfo info)
        {
            pckagentinfo = info;
            docagentMap = new Dictionary<string, DocAgent>();
        }

    }
}
