using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Icris.Excel2Api.Console
{
    class Program
    {
        static void Main(string[] args)
        {
            Microsoft.Owin.Hosting.WebApp.Start<Startup>(new Microsoft.Owin.Hosting.StartOptions("http://localhost:7092"));
            
            
            System.Console.WriteLine("Up & running at http://localhost:7092/app/index.html, press enter to exit");
            System.Console.ReadLine();

        }
    }
}
