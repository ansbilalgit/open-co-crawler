using OpenCoCrawler.Extensions;
using OpenCoCrawler.Utilitis;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static OpenCoCrawler.Infrastructure.ExcelStructs;

namespace OpenCoCrawler
{
    public class ExcelManager
    {
        private static string filePath = ConfigurationManager.AppSettings["ExcelFilePath"];
        private static string saleRecordSheet = ConfigurationManager.AppSettings["SalesRecordSheetName"];
        private static string companiesRecordSheet = ConfigurationManager.AppSettings["CompaniesRecordSheetName"];
        private static string OutputFilePath = ConfigurationManager.AppSettings["OutputFilePath"];

        public static void SortSalesRecords()
        {
            try
            {
                var dataSet = ExcelUtilities.GetDataSetFromExcelFile(filePath, saleRecordSheet);
                var salesTable = dataSet.Tables[saleRecordSheet];
                var sortKey = "Transfer Date";
                var sortedTable = SortDataTable(salesTable, sortKey);

                ExcelUtilities.ExportToExcelOleDb(sortedTable, saleRecordSheet, OutputFilePath, "Sorted.xlsx", true);

            }
            catch (Exception ex)
            {

                LoggerUtility.Write("Exception in SortSaleRecords : ", ex.Message);
            }

        }

        private static DataTable SortDataTable(DataTable dt1, string key)
        {
            // have data
            var startDate = new DateTime(2016, 1, 1).ToShortDateString();
            var endDate = new DateTime(2020, 1, 1).ToShortDateString();
            DataTable dt2 = new DataTable(); // temp data table
            DataRow[] dra = dt1.Select("[" + key + "] >= #" + startDate + "#  AND [" + key + "] <= #" + endDate + "#", "[" + key + "] DESC");
            if (dra.Length > 0)
                dt2 = dra.CopyToDataTable();
            else
                dt2 = dt1;
            return dt2;
        }


    }
}
