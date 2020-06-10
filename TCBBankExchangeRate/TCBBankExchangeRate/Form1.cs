using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Net;
using System.Xml;
using System.Windows.Forms;

namespace TCBBankExchangeRate
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            Crawler();
        }

        public void Crawler()
        {
            string url = "https://www.tcb-bank.com.tw/finance_info/Pages/foreign_spot_rate.aspx";

            WebClient client = new WebClient();
            client.Encoding = Encoding.UTF8;
            // 無法建立 SSL / TLS 的安全通道。解決方法 [For .NET Framwork 4.5+]
            System.Net.ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
            System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls;
            string html= client.DownloadString(url);

            //尋找 id="ctl00_PlaceHolderEmptyMain_PlaceHolderMain_fecurrentid_gvResult" 
            //	<table class="table" cellspacing="0" rules="all" border="1" id="ctl00_PlaceHolderEmptyMain_PlaceHolderMain_fecurrentid_gvResult" width="100%">
            int idx = html.IndexOf("ctl00_PlaceHolderEmptyMain_PlaceHolderMain_fecurrentid_gvResult");
            int idxTableStart = html.LastIndexOf("<table",idx);
            int idxTableEnd = html.IndexOf("</table>", idx);
            string TableXml = html.Substring(idxTableStart, idxTableEnd - idxTableStart + "</table>".Length);

            XmlDocument XmlDoc = new XmlDocument();
            XmlDoc.LoadXml(TableXml);
            XmlNode CaptionNode = XmlDoc.SelectSingleNode("table/caption");
            this.Text = CaptionNode.InnerText + " [" + url + "]";
            XmlNodeList NodeLists = XmlDoc.SelectNodes("table/tr");
            DataTable dt = new DataTable();
            //設定Header
            for (int i = 0; i < NodeLists[0].ChildNodes.Count; i++)
            {
                dt.Columns.Add(NodeLists[0].ChildNodes[i].InnerText, typeof(String));
            }
            //擷取匯率
            for (int i = 1; i < NodeLists.Count - 1; i++)
            {
                DataRow row = dt.NewRow();
                for (int j =0; j < NodeLists[i].ChildNodes.Count; j++)
                {
                    row[j] = NodeLists[i].ChildNodes[j].InnerText;
                }
                dt.Rows.Add(row);
            }
            dataGridView1.DataSource = dt;
            
        }
    }
}
