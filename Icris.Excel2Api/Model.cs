using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Icris.Excel2Api
{
    public class Model
    {
        public Model()
        {
            Inputs = new Dictionary<string, Input>();
            Outputs = new Dictionary<string, Output>();
        }
        public Dictionary<string,Input> Inputs { get; set; }
        public Dictionary<string,Output> Outputs { get; set; }
    }
}
