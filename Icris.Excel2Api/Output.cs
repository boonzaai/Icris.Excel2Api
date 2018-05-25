using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Icris.Excel2Api
{
    public class Output
    {
        public string Name { get; set; }
        public string Description { get; set; }
        public string Unit { get; set; }
        public object Value { get; set; }
        public int Row { get; set; }
    }
}
