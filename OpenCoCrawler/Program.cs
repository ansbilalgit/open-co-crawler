using OpenCoCrawler.Utilitis;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenCoCrawler
{
    class Program
    {
        static void Main(string[] args)
        {
            //ExcelManager.testRead();
            int i = 0;
            while (i == 0)
            {
                Console.WriteLine("Enter '1' to Filter Sales Record");
                Console.WriteLine("Enter '2' to Crawl Companies Data");
                Console.WriteLine("Enter '3' to Exit");
                string input = Console.ReadLine();

                int.TryParse(input, out i);
                if (i == 0)
                {
                    Console.WriteLine("Invalid Input");
                    Console.WriteLine("___________________________________________");
                }
                else
                if (i == 1)
                {
                    Console.WriteLine("Sales Record Filtering Started...");
                    ExcelManager.SortSalesRecords();
                    i = 0;
                }
                else if (i == 2)
                {
                    Console.WriteLine("Company Data Crawling Started...");
                    ExcelManager.FilterCompaniesRecords();
                    i = 0;
                }
                else if (i == 3)
                    break;
                else
                    i = 0;

            }


            ////ExcelManager.SortSalesRecords();
            //ExcelManager.FilterCompaniesRecords();
            ////ProxyManager pm = new ProxyManager();

            //Console.ReadKey();
        }
    }
}
