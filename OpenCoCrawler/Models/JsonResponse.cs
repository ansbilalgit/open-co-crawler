using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenCoCrawler.Models
{
    public class JsonResponse
    {

        public string api_version { get; set; }
        public SearchResult results { get; set; }
    }

    public class JsonResponseCompanyResult
    {

        public string api_version { get; set; }
        public CompanyResult results { get; set; }
    }

    public class CompanyResult
    {
        public Company company { get; set; }
        public int page { get; set; }
        public int per_page { get; set; }
        public int total_pages { get; set; }
        public int total_count { get; set; }
    }

    public class SearchResult
    {
        public List<CompanyObject> companies { get; set; }
        public int page { get; set; }
        public int per_page { get; set; }
        public int total_pages { get; set; }
        public int total_count { get; set; }
    }



    public class Source
    {
        public string publisher { get; set; }
        public string url { get; set; }
        public DateTime retrieved_at { get; set; }
    }

    public class RegisteredAddress
    {
        public string street_address { get; set; }
        public string locality { get; set; }
        public string region { get; set; }
        public string postal_code { get; set; }
        public string country { get; set; }
    }

    public class CompanyObject
    {

        public Company company { get; set; }
    }

    public class Company
    {
        public string name { get; set; }
        public string company_number { get; set; }
        public string jurisdiction_code { get; set; }
        public string incorporation_date { get; set; }
        public object dissolution_date { get; set; }
        public string company_type { get; set; }
        public string registry_url { get; set; }
        public object branch { get; set; }
        public object branch_status { get; set; }
        public bool? inactive { get; set; }
        public string current_status { get; set; }
        public DateTime created_at { get; set; }
        public DateTime updated_at { get; set; }
        public DateTime retrieved_at { get; set; }
        public string opencorporates_url { get; set; }
        public Source source { get; set; }
        public string agent_name { get; set; }
        //public RegisteredAddress agent_address { get; set; }
        public object number_of_employees { get; set; }
        public object native_company_number { get; set; }
        public string registered_address_in_full { get; set; }
        public RegisteredAddress registered_address { get; set; }
        public Data data { get; set; }
        public object financial_summary { get; set; }
        public object home_company { get; set; }
        public object controlling_entity { get; set; }
        public List<object> ultimate_beneficial_owners { get; set; }
        public List<Officer> officers { get; set; }
    }

    public class Datum
    {
        public int id { get; set; }
        public string title { get; set; }
        public string data_type { get; set; }
        public string description { get; set; }
        public string opencorporates_url { get; set; }
    }
    public class Officer
    {
        public Officer2 officer { get; set; }
    }

    public class Officer2
    {
        public int id { get; set; }
        public string name { get; set; }
        public string position { get; set; }
        public string uid { get; set; }
        public object start_date { get; set; }
        public object end_date { get; set; }
        public string opencorporates_url { get; set; }
        public object occupation { get; set; }
        public bool inactive { get; set; }
        public object current_status { get; set; }
    }

    public class Filing2
    {
        public int id { get; set; }
        public string title { get; set; }
        public string uid { get; set; }
        public string opencorporates_url { get; set; }
        public string date { get; set; }
    }

    public class Filing
    {
        public Filing2 filing { get; set; }
    }

    public class MostRecent
    {
        public Datum datum { get; set; }
    }

    public class Data
    {
        public List<MostRecent> most_recent { get; set; }
        public int total_count { get; set; }
        public string url { get; set; }
    }



}
