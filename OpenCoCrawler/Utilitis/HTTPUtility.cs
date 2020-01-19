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
        private static ProxyManager pm = new ProxyManager();
        private static List<ProxyModel> proxyList = pm.GetProxyList();
        private static int count = 0;

        public static string getStringFromUrl(string Url)
        {
            string json = string.Empty;
            try
            {
                if (count < proxyList.Count())
                {
                    using (WebClient wc = new WebClient())
                    {
                        ProxyModel p = proxyList[count];
                        WebProxy webProxy = new WebProxy(p.IP, p.Port);

                        wc.Proxy = webProxy;
                        var str = wc.DownloadString(Url);

                        json = JObject.Parse(str).ToString();
                    }
                }
                else
                {
                    _updateProxyList();
                }
            }
            catch (Exception ex)

            {

                LoggerUtility.Write(ex.Message, Url);
                count++;
                if (count >= proxyList.Count)
                    _updateProxyList();

                LoggerUtility.Write("Changing IP to ", proxyList[count].IP);
                return getStringFromUrl(Url);
            }
            return json;
        }

        private static void _updateProxyList()
        {
            LoggerUtility.Write("Proxy List Ended", "Updating List");
            proxyList = pm.UpdateProxyList();
            count = 0;

        }
    }
}
