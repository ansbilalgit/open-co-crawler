using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace OpenCoCrawler.Utilitis
{
    public class HTTPUtility
    {
        public static string getStringFromUrl(string Url)
        {
            string json = string.Empty;
            try
            {
                using (WebClient wc = new WebClient())
                {
                    //WebProxy webProxy = new WebProxy("http://154.0.233:8080/");
                    //wc.Proxy = webProxy;
                    var str = wc.DownloadString(Url);

                    json = JObject.Parse(str).ToString();
                }
            }
            catch (Exception ex)

            {

                LoggerUtility.Write(ex.Message, Url);
            }
            return json;
        }
    }
}
