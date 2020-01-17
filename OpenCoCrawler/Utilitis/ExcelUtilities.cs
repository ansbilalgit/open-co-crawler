using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static OpenCoCrawler.Infrastructure.ExcelStructs;

namespace OpenCoCrawler.Utilitis
{
    public class ExcelUtilities
    {
        public static DataSet GetDataSetFromExcelFile(string file, string sheetName)
        {
            try
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
            catch (Exception)
            {

                throw;
            }
        }



        public static void ExportToExcelOleDb(DataTable dataTable, string tableName, string connectionString,
                                                          string fileName, bool deleteExistFile)
        {
            // Support for existing file overwrite.
            string fn = connectionString + fileName;
            if (deleteExistFile && File.Exists(fn))
                File.Delete(fn);
            ExportToExcelOleDb(dataTable, tableName, connectionString, fileName);
        }

        private static bool ExportToExcelOleDb(DataTable dataTable, string tableName, string outputFilePath, string fileName)
        {
            try
            {
                string connectionString = GetConnectionString(outputFilePath + fileName);
                // Check for null set.
                if (dataTable != null && dataTable.Rows.Count > 0)
                {
                    using (OleDbConnection connection = new OleDbConnection(connectionString))
                    {
                        // Initialise SqlCommand and open.
                        OleDbCommand command = null;
                        connection.Open();


                        // Build the Excel create table command.
                        string strCreateTableStruct = BuildCreateTableCommand(dataTable, tableName);
                        if (String.IsNullOrEmpty(strCreateTableStruct))
                            return false;
                        command = new OleDbCommand(strCreateTableStruct, connection);
                        command.ExecuteNonQuery();

                        // Puch each row into Excel.
                        for (int rowIndex = 0; rowIndex < dataTable.Rows.Count; rowIndex++)
                        {
                            command = new OleDbCommand(BuildInsertCommand(dataTable, tableName, rowIndex), connection);
                            command.ExecuteNonQuery();


                        }

                        connection.Close();
                    }
                }
                return true;
            }
            catch (Exception eX)
            {
                return false;
            }
        }


        private static string[] BuildExcelSheetNames(string connectionString)
        {
            // Variables.
            DataTable dt = null;
            string[] excelSheets = null;

            using (OleDbConnection schemaConn = new OleDbConnection(connectionString))
            {
                schemaConn.Open();
                dt = schemaConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                // No schema found.
                if (dt == null)
                    return null;

                // Insert 'TABLE_NAME' to sheet name array.
                int i = 0;
                excelSheets = new string[dt.Rows.Count];
                foreach (DataRow row in dt.Rows)
                    excelSheets[i++] = row["TABLE_NAME"].ToString();
            }
            return excelSheets;
        }


        private static string BuildCreateTableCommand(DataTable dataTable, string tabeName)
        {
            // Get the type look-up tables.
            StringBuilder sb = new StringBuilder();
            Dictionary<string, string> dataTypeList = BuildExcelDataTypes();

            // Check for null data set.
            if (dataTable.Columns.Count <= 0)
                return null;

            // Start the command build.
            sb.AppendFormat("CREATE TABLE [{0}] (", tabeName);

            // Build column names and types.
            foreach (DataColumn col in dataTable.Columns)
            {
                string type = ExcelDataTypes.TEXT;
                if (dataTypeList.ContainsKey(col.DataType.Name.ToString().ToLower()))
                {
                    type = dataTypeList[col.DataType.Name.ToString().ToLower()];
                }
                sb.AppendFormat("[{0}] {1},", col.Caption.Replace(' ', '_'), type);
            }
            sb = sb.Replace(',', ')', sb.ToString().LastIndexOf(','), 1);
            return sb.ToString();
        }


        private static string BuildInsertCommand(DataTable dataTable, string tableName, int rowIndex)
        {
            StringBuilder sb = new StringBuilder();

            // Remove whitespace.
            sb.AppendFormat("INSERT INTO [{0}](", tableName);
            foreach (DataColumn col in dataTable.Columns)
                sb.AppendFormat("[{0}],", col.Caption.Replace(' ', '_'));
            sb = sb.Replace(',', ')', sb.ToString().LastIndexOf(','), 1);

            // Write values.
            sb.Append("VALUES (");
            foreach (DataColumn col in dataTable.Columns)
            {
                string type = col.DataType.ToString();
                string strToInsert = String.Empty;
                strToInsert = dataTable.Rows[rowIndex][col].ToString().Replace("'", "''");
                sb.AppendFormat("'{0}',", strToInsert);
                //strToInsert = String.IsNullOrEmpty(strToInsert) ? "NULL" : strToInsert;
                //String.IsNullOrEmpty(strToInsert) ? "NULL" : strToInsert);
            }
            sb = sb.Replace(',', ')', sb.ToString().LastIndexOf(','), 1);
            return sb.ToString();
        }


        private static string BuildExcelSheetName(DataTable dataTable)
        {
            string retVal = dataTable.TableName;
            //if (dataTable.ExtendedProperties.ContainsKey(TABLE_NAME_PROPERTY))
            //    retVal = dataTable.ExtendedProperties[TABLE_NAME_PROPERTY].ToString();
            //return retVal.Replace(' ', '_');
            return retVal;
        }

        /// <summary>
        /// Dictionary for conversion between .NET data types and Excel 
        /// data types. The conversion does not currently work, so I am 
        /// pushing all data upto excel as Excel "TEXT" type.
        /// </summary>
        /// <returns></returns>
        private static Dictionary<string, string> BuildExcelDataTypes()
        {
            Dictionary<string, string> dataTypeLookUp = new Dictionary<string, string>();

            // I cannot get the Excel formatting correct here!?
            dataTypeLookUp.Add(NETDataTypes.SHORT, ExcelDataTypes.TEXT);
            dataTypeLookUp.Add(NETDataTypes.INT, ExcelDataTypes.TEXT);
            dataTypeLookUp.Add(NETDataTypes.LONG, ExcelDataTypes.TEXT);
            dataTypeLookUp.Add(NETDataTypes.STRING, ExcelDataTypes.TEXT);
            dataTypeLookUp.Add(NETDataTypes.DATE, ExcelDataTypes.TEXT);
            dataTypeLookUp.Add(NETDataTypes.BOOL, ExcelDataTypes.TEXT);
            dataTypeLookUp.Add(NETDataTypes.DECIMAL, ExcelDataTypes.TEXT);
            dataTypeLookUp.Add(NETDataTypes.DOUBLE, ExcelDataTypes.TEXT);
            dataTypeLookUp.Add(NETDataTypes.FLOAT, ExcelDataTypes.TEXT);
            return dataTypeLookUp;
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

    }
}
