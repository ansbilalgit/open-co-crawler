using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenCoCrawler.Utilitis
{
    public class ProxyManager
    {
        private List<ProxyModel> proxyList = new List<ProxyModel>();

        public ProxyManager()
        {
            this.proxyList = updateProxyList();
        }

        public List<ProxyModel> GetProxyList()
        {
            return this.proxyList;
        }

        public List<ProxyModel> UpdateProxyList()
        {
            this.proxyList = updateProxyList();
            return this.proxyList;
        }

        private List<ProxyModel> updateProxyList()
        {
            List<ProxyModel> pl = new List<ProxyModel>();
            HtmlWeb web = new HtmlWeb();
            HtmlDocument doc = web.Load("https://free-proxy-list.net/");
            var table = doc.DocumentNode.SelectSingleNode("//table[@id='proxylisttable']");
            if (table != null)
            {
                var tbody = table.Descendants("tbody").FirstOrDefault();
                if (tbody != null)
                {
                    var rowList = tbody.Descendants("tr");
                    foreach (var tr in rowList)
                    {
                        var tdList = tr.Descendants("td").ToList();
                        if (tdList != null && tdList.Count() > 1)
                        {
                            string ip = tdList[0].InnerText;
                            int _port = 8080;

                            string port = tdList[1].InnerText;
                            int.TryParse(port, out _port);
                            ProxyModel p = new ProxyModel();
                            p.IP = ip;
                            p.Port = _port;
                            pl.Add(p);
                        }
                    }
                }
            }
            return pl;
        }
    }

    public class ProxyModel
    {
        public string IP { get; set; }
        public int Port { get; set; }
    }
}
