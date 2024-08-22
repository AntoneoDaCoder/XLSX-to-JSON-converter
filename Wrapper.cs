using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_to_JSON_converter
{
    internal class Wrapper
    {
        public PckAgent pckagent { get; set; }
        public Wrapper(PckAgent agent) { 
            pckagent = agent;
        }
    }
}
