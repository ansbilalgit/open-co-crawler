using Newtonsoft.Json;
using OpenCoCrawler.Enums;
using OpenCoCrawler.Models;
using OpenCoCrawler.Utilitis;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OpenCoCrawler
{
    public class APIManager
    {
        public static DataTable ParseCompaniesData(DataTable filteredTable)
        {
            DataTable updatedTable = new DataTable();
            try
            {
                updatedTable = filteredTable.Clone();
                int i = 0;
                foreach (DataRow row in filteredTable.Rows)
                {
                    bool isCompanyCrawled = false;
                    DataRow dr = crawlRow(row, out isCompanyCrawled);
                    DataRow _dr = updatedTable.NewRow();
                    updatedTable.ImportRow(dr);
                    i++;
                    if (isCompanyCrawled)
                        Console.WriteLine("Company Data Found :" + row[CompanyEnum.COMPANY_NAME]);
                }

            }
            catch (Exception ex)
            {
                LoggerUtility.Write("Data Parsing Failed ", ex.Message);
            }
            return updatedTable;
        }

        private static DataRow crawlRow(DataRow row, out bool isCompanyCrawled)
        {
            DataRow dr = row;
            isCompanyCrawled = false;
            try

            {
                string t = " A J Potter Investments, Llc";
                string searchQuery = "http://api.opencorporates.com/v0.4/companies/search?q=" + dr[CompanyEnum.COMPANY_NAME].ToString().Replace(' ', '+');
                string companyListString = HTTPUtility.getStringFromUrl(searchQuery);

                var jsonResponse = JsonConvert.DeserializeObject<JsonResponse>(companyListString);
                if (jsonResponse != null && jsonResponse.results.companies.Count > 0)
                {

                    var selectedCompany = jsonResponse.results.companies.FirstOrDefault(c => c.company.inactive == false).company;
                    if (selectedCompany != null)
                    {

                        Thread.Sleep(3000);
                        string coQuery = "http://api.opencorporates.com/v0.4/companies/" + selectedCompany.jurisdiction_code + "/" + selectedCompany.company_number;

                        string companyResultString = HTTPUtility.getStringFromUrl(coQuery);
                        Thread.Sleep(3000);
                        var companyResultResponse = JsonConvert.DeserializeObject<JsonResponseCompanyResult>(companyResultString);
                        var companyData = companyResultResponse.results.company;

                        dr[CompanyEnum.OWNER_F_NAME] = getName(companyData.agent_name, true);
                        dr[CompanyEnum.OWNER_l_NAME] = getName(companyData.agent_name, false);
                        dr[CompanyEnum.OWNER_ADDR] = companyData.registered_address != null ? companyData.registered_address.street_address : companyData.registered_address_in_full;
                        dr[CompanyEnum.OWNER_CITY] = companyData.registered_address?.locality ?? "";
                        int z = 0000;
                        int.TryParse(companyData.registered_address?.postal_code, out z);
                        dr[CompanyEnum.OWNER_STATE] = companyData.registered_address?.region ?? "";
                        dr[CompanyEnum.OWNER_ZIP] = z;
                        string mailDescription = string.IsNullOrEmpty(getMailAddr(companyData)) ? companyData.registered_address_in_full ?? "" : getMailAddr(companyData);
                        dr[CompanyEnum.MAIL_ADDR] = mailDescription;
                        dr[CompanyEnum.MAIL_UNIT] = getMailUnits(mailDescription, 1);
                        dr[CompanyEnum.MAIL_CITY] = getMailUnits(mailDescription, 2);
                        dr[CompanyEnum.MAIL_STATE] = getMailUnits(mailDescription, 3);
                        dr[CompanyEnum.MAIL_ZIP] = getMailUnits(mailDescription, 4);

                        isCompanyCrawled = true;
                    }

                }

            }
            catch (Exception ex)
            {
                LoggerUtility.Write("Failed to crawl " + row["company name"].ToString(), ex.Message);
                isCompanyCrawled = false;
            }
            return dr; ;
        }

        private static int? GetPostalCode(string postal_code)
        {
            int p = 0000;
            if (!string.IsNullOrEmpty(postal_code) && postal_code.Contains('-'))
            {
                var token = postal_code.Split('-');
                int.TryParse(token[0], out p);
            }
            else
            {
                int.TryParse(postal_code, out p);
            }
            return p;
        }

        private static int? getZipCode(string postal_code)
        {
            int z = 0000;
            //if(postal_code.Contains)
            return z;
        }

        private static string getMailUnits(string mailDescription, int addrType)
        {
            string u = string.Empty;
            var addrToken = mailDescription.Split(',');
            if (addrType == 1)
            {
                if (addrToken.Length >= 2)
                    u = addrToken[0] + addrToken[1];
                else u = addrToken[0];
            }
            if (addrType == 2)
            {
                if (addrToken.Length - 3 >= 0)
                    u = addrToken[addrToken.Length - 3];
                else u = addrToken[0];
            }
            if (addrType == 3)
            {
                if (addrToken.Length - 2 >= 0)
                    u = addrToken[addrToken.Length - 2];
                else u = addrToken[0];
            }
            if (addrType == 4)
            {
                if (addrToken.Length - 1 >= 0)
                    u = addrToken[addrToken.Length - 1];
                else u = addrToken[0];
            }
            return u;
        }

        private static string getMailAddr(Company companyData)
        {
            string mail = string.Empty;

            if (companyData.data != null)
            {
                var mostRecent = companyData.data.most_recent;
                if (mostRecent != null)
                {

                    var datum = mostRecent.FirstOrDefault(mr => mr.datum != null && mr.datum.title == CompanyEnum.MAILING_ADDR).datum;
                    if (datum != null)
                    {
                        mail = datum.description;
                    }
                }
            }
            return mail;

        }

        private static string getName(string agent_name, bool isFirst)
        {
            string name = agent_name;
            if (agent_name != null)
            {
                var nameTokens = agent_name.Contains(',') ? agent_name.Split(',') : agent_name.Contains('.') ? agent_name.Split('.') : agent_name.Split(' ');
                if (isFirst)
                {
                    if (nameTokens.Length >= 2)
                        name = nameTokens[1];
                }
                else name = nameTokens[0];
            }

            return name;
        }
    }
}
