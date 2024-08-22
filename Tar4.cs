using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_to_JSON_converter
{
    internal class Tar4
    {
        public byte nmonth { get; set; }
        public double nsummonth { get; set; }
        public LinkedList<Dictionary<string, double>> tar4sum { get; set; }
        public Tar4()
        {
            tar4sum = new LinkedList<Dictionary<string, double>>();
        }
    }
}
