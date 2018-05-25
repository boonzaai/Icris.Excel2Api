using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Icris.Excel2Api
{
    public class Input
    {
        public Input()
        {
            Options = new List<string>();
        }
        public string Name { get; set; }
        public string Description { get; set; }
        public string Unit { get; set; }
        public object Value { get; set; }
        public string Errormessage { get; set; }
        public List<string> Options { get; set; }
        public bool Enabled { get; set; }
        public bool Valid { get; set; }
        public int Row { get; set; }
    }
}
