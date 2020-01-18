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
        #region Private Prop
        private static string filePath = ConfigurationManager.AppSettings["ExcelFilePath"];
        private static string saleRecordSheet = ConfigurationManager.AppSettings["SalesRecordSheetName"];
        private static string companiesRecordSheet = ConfigurationManager.AppSettings["CompaniesRecordSheetName"];
        private static string OutputFilePath = ConfigurationManager.AppSettings["OutputFilePath"];
        private static string CompanyTypes = ConfigurationManager.AppSettings["CompanyTypes"];
        #endregion

        public static void SortSalesRecords()
        {
            try
            {
                var dataSet = ExcelUtilities.GetDataSetFromExcelFile(filePath, saleRecordSheet);
                var salesTable = dataSet.Tables[saleRecordSheet];
                var sortKey = "Transfer Date";
                var sortedTable = SortDataTable(salesTable, sortKey);
                string outputFileName = "SortedSalesRecords.xlsx";
                ExcelUtilities.ExportToExcelOleDb(sortedTable, saleRecordSheet, OutputFilePath, outputFileName, true);

                LoggerUtility.Write("Success", "Sales Record Sorted " + OutputFilePath + outputFileName);

                
            }
            catch (Exception ex)
            {

                LoggerUtility.Write("Exception in SortSaleRecords : ", ex.Message);
            }

        }

        public static void FilterCompaniesRecords()
        {
            var dataSet = ExcelUtilities.GetDataSetFromExcelFile(filePath, companiesRecordSheet);
            var companiesTable = dataSet.Tables[companiesRecordSheet];
            var filteredTable = FilterDataTable(companiesTable);
            string outputFileName = "FilteredCompanyRecords.xlsx";
            ExcelUtilities.ExportToExcelOleDb(filteredTable, companiesRecordSheet, OutputFilePath, outputFileName, true);

            LoggerUtility.Write("Success", "Company Records Sorted " + OutputFilePath + outputFileName);
            DataTable updatedResults = APIManager.ParseCompaniesData(filteredTable);

            ExcelUtilities.ExportToExcelOleDb(filteredTable, companiesRecordSheet, OutputFilePath, "updatedCompanies.xlsx", true);


        }
        private static DataTable FilterDataTable(DataTable dt1)
        {
            DataTable dt2 = new DataTable();
            try
            {
                dt2 = dt1.Clone();
                foreach (DataRow row in dt1.Rows)
                {
                    if (isValidCompany(row["company name"].ToString()))
                    {
                        DataRow dr = dt2.NewRow();
                        dt2.ImportRow(row);
                    }
                }

            }
            catch (Exception ex)
            {

                LoggerUtility.Write("Filtering Company list", ex.Message);
            }
            return dt2;
        }

        private static bool isValidCompany(string n)
        {
            bool result = false;
            var validTags = CompanyTypes.Split(',');
            foreach (var tag in validTags)
            {
                if (n.ToLower().EndsWith(tag.ToLower()))
                {
                    result = true;
                    break;
                }
            }
            return result;
        }



        #region Private Methods

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
        #endregion


    }
}
