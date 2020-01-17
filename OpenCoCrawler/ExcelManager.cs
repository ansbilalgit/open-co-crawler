using OpenCoCrawler.Extensions;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenCoCrawler
{
    public class ExcelManager
    {
        private static string filePath = ConfigurationManager.AppSettings["ExcelFilePath"];
        private static string saleRecordSheet = ConfigurationManager.AppSettings["SalesRecordSheetName"];
        private static string companiesRecordSheet = ConfigurationManager.AppSettings["CompaniesRecordSheetName"];
        private static string OutputFilePath = ConfigurationManager.AppSettings["OutputFilePath"];
        public static void testRead()
        {


            var dataSet = GetDataSetFromExcelFile(filePath, saleRecordSheet);

            //Console.WriteLine(string.Format("reading file: {0}", file));
            //Console.WriteLine(string.Format("coloums: {0}", dataSet.Tables[0].Columns.Count));
            //Console.WriteLine(string.Format("rows: {0}", dataSet.Tables[0].Rows.Count));
            Console.ReadKey();
        }

        public static void SortSalesRecords()
        {
            var dataSet = GetDataSetFromExcelFile(filePath, saleRecordSheet);
            var salesTable = dataSet.Tables[saleRecordSheet];
            var sortKey = "Transfer Date";
            var sortedTable = SortDataTable(salesTable, sortKey);
            var salesData = sortedTable.ToDynamic();


        }

        private static DataTable SortDataTable(DataTable dt1, string key)
        {
            // have data
            var startDate = new DateTime(2016, 1, 1).ToString("M/d/yyyy");
            var endDate = new DateTime(2019, 1, 1).ToString("M/d/yyyy");
            DataTable dt2 = new DataTable(); // temp data table
            DataRow[] dra = dt1.Select("[" + key + "] >= #" + startDate + "#  AND [" + key + "] <= #" + endDate + "#", "[" + key + "] DESC");
            if (dra.Length > 0)
                dt2 = dra.CopyToDataTable();
            else
                dt2 = dt1;
            return dt2;
        }

        private static List<dynamic> Sort(List<dynamic> input, string property)
        {
            return input.OrderBy(p => p.GetType()
                                       .GetProperty(property)
                                       .GetValue(p, null)).ToList();
        }

        private static string GetConnectionString(string file)
        {
            Dictionary<string, string> props = new Dictionary<string, string>();

            string extension = file.Split('.').Last();

            if (extension == "xls")
            {
                //Excel 2003 and Older
                props["Provider"] = "Microsoft.Jet.OLEDB.4.0";
                props["Extended Properties"] = "Excel 8.0";
            }
            else if (extension == "xlsx")
            {
                //Excel 2007, 2010, 2012, 2013
                props["Provider"] = "Microsoft.ACE.OLEDB.12.0;";
                props["Extended Properties"] = "Excel 12.0 XML";
            }
            else
                throw new Exception(string.Format("error file: {0}", file));

            props["Data Source"] = file;

            StringBuilder sb = new StringBuilder();

            foreach (KeyValuePair<string, string> prop in props)
            {
                sb.Append(prop.Key);
                sb.Append('=');
                sb.Append(prop.Value);
                sb.Append(';');
            }

            return sb.ToString();
        }

        private static DataSet GetDataSetFromExcelFile(string file, string sheetName)
        {
            DataSet ds = new DataSet();

            string connectionString = GetConnectionString(file);

            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = conn;




                // Get all rows from the Sheet
                cmd.CommandText = "SELECT * FROM [" + sheetName + "$]";

                DataTable dt = new DataTable();
                dt.TableName = sheetName;

                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);

                ds.Tables.Add(dt);


                cmd = null;
                conn.Close();
            }

            return ds;
        }

    }
}
