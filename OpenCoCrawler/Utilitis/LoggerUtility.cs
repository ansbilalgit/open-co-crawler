using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenCoCrawler.Utilitis
{
    public class LoggerUtility
    {
        public static void Write(string info, string m)
        {
            Console.WriteLine("___________________________________________");
            Console.WriteLine("Info : " + info);
            Console.WriteLine("Log Message : " + m);
            Console.WriteLine("____________________________________________");
        }
    }
}
