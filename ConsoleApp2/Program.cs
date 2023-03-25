using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;
using UglyToad.PdfPig;
using System.Text.RegularExpressions;
using System.Web;

namespace ConsoleApp2
{
    internal class Program
    {
        public static HttpClient client { get; set; }
        public static CookieContainer cookie { get; set; }
        public static HttpClientHandler handler { get; set; }
        public static string addingbody="---Remark:---\n";
        static void Main(string[] args)
        {
            string[] PTSvalue = Console.ReadLine().Split(';');
            System.Net.ServicePointManager.Expect100Continue = false;
            cookie = new CookieContainer();
            handler = new HttpClientHandler();
            handler.CookieContainer = cookie;
            client = new HttpClient(handler);
            client.DefaultRequestHeaders.Add("Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9");
            client.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/103.0.0.0 Safari/537.36");
            Autherzire();
            foreach (var item in PTSvalue)
            {
                string g_pts = mgss($"https://network.infornexus.com/en/trade/PlantoShipFolder?key={item}").Content.ReadAsStringAsync().Result;
                var qtitle = g_pts.Split('<').Where(INV => INV.Contains("Packing")).Where(INV => INV.Contains("List")).Where(INV => INV.Contains("listtablecell"));
                int qnumber = Enumerable.Range(0, g_pts.Split('<').Length).Where(i => g_pts.Split('<')[i] == qtitle.ToArray()[0]).ToArray()[0];
                string linkPKL = "https://network.infornexus.com/en/trade/" + g_pts.Split('<')[qnumber-5].Split('\"')[1];
                string pkllineget = mgss(linkPKL).Content.ReadAsStringAsync().Result.Split('<').Where(iv => iv.Contains("loadPDF")).ToArray()[0].Split('\"')[3];
                byte[] bytesPTS = mgss($"https://network.infornexus.com/dyncon/?producer=PlatformTemplateProducer&topicName=VendorBookingRequest_viewPdf&rootId={(g_pts.Split('\"').Where(INV => INV.Contains("VendorBookingRequest?key")).ToArray()[0]).Split('=')[1]}&pmId=-1047&renderType=PDF&type=VendorBookingRequest&isHuman=true").Content.ReadAsByteArrayAsync().Result;
                Read_PTSfile(bytesPTS);
                byte[] bytesPACKING = mgss($"https://network.infornexus.com/en/trade/PackingManifestPDF.jsp?key={pkllineget}&OrderAssignment=&OrderAssignmentTYPE=").Content.ReadAsByteArrayAsync().Result;
                Read_PKLfile(bytesPACKING,item);
                File.WriteAllBytes($"{Directory.GetCurrentDirectory()}\\PTS # {item} _ PKL# {pkllineget} PackingManifestPDF .pdf", bytesPACKING);
                Console.WriteLine($"Done Loading PackingManifestPDF and save file {item}");
            }
            Console.WriteLine("Done export file");
            string[] attachfiles = find_file_in_path("PackingManifestPDF");
            Console.WriteLine(addingbody);
            SendEmail($"PACKING LIST _ CREATE DATE {DateTime.Now.Date.ToString()}", $"Dear BU Team ,\n<br> Pls file PackingManifestPDF file in the attach <br>\nThank you! \n<br><br>{addingbody}<br><br>--Ai02--", File.ReadAllText(Directory.GetCurrentDirectory() + "\\to.txt"), File.ReadAllText(Directory.GetCurrentDirectory() + "\\cc.txt"), attachfiles);
            Console.ReadKey();
        }
        public class PTSPO
        {
            public string PO { get; set; }
            public string qty { get; set; }
            public string carton { get; set; }
            public PTSPO(string PO, string qty, string carton)
            {
                this.PO = PO;
                this.qty = qty;
                this.carton = carton;
            }
        }
        public static void Read_PKLfile(byte[] bytes,string pts)
        {
            using (PdfDocument document = PdfDocument.Open(bytes))
            {
                var lingw = document.GetPages().ToArray()[0];
                var array = lingw.GetWords().ToArray();
                int start = Array.FindIndex(array, itemc => itemc.Text.Contains("Unit")) + 6;
                int end = Array.FindIndex(array, itemc => itemc.Text.Contains("Totals"));
                string valv = "";int check = 0;
                while (true)
                {
                    if(int.TryParse(array[start].Text, out check) && array[start - 1].Text == "Unit")
                    {
                        break;
                    }
                    start++;
                }
                for (int i = start; i < end; i++)
                {
                    valv = valv + array[i] + "\t";
                    if ((i - start + 1) % 12 == 0)
                    {
                        valv = valv + "\n";
                    }
                }

                foreach (string v in valv.Split('\n'))
                {
                    if (v != "")
                    {
                        if (ps.Where(ivn => ivn.PO == v.Split('\t')[0]).ToArray()[0].qty == v.Split('\t')[2] && ps.Where(ivn => ivn.PO == v.Split('\t')[0]).ToArray()[0].carton == v.Split('\t')[3])
                        {
                            Console.WriteLine($"PTS# {pts}: \nPO# {v.Split('\t')[0]} _Qty: {v.Split('\t')[2]} _Carton: {v.Split('\t')[3]} _ PASS TEST PTS SAME PACKING LIST");
                            addingbody = addingbody + $"<br>\nPTS# {pts}: \nPO# {v.Split('\t')[0]} _Qty: {v.Split('\t')[2]} _Carton: {v.Split('\t')[3]} _ PASS TEST PTS SAME PACKING LIST\n";
                        }
                        else
                        {
                            Console.WriteLine($"PTS# {pts}:\nPKL file show PO# {v.Split('\t')[0]} _Qty: {v.Split('\t')[2]} _Carton: {v.Split('\t')[3]}");
                            Console.WriteLine($"PTS file show PO# {v.Split('\t')[0]} _Qty: {ps.Where(ivn => ivn.PO == v.Split('\t')[0]).ToArray()[0].qty} _Carton: {ps.Where(ivn => ivn.PO == v.Split('\t')[0]).ToArray()[0].carton}");
                            addingbody = addingbody + $"PTS# {pts}:\nPKL GTN file show PO# {v.Split('\t')[0]} _Qty: {v.Split('\t')[2]} _Carton: {v.Split('\t')[3]}" + $"\nBut PTS file show PO# {v.Split('\t')[0]} _Qty: {ps.Where(ivn => ivn.PO == v.Split('\t')[0]).ToArray()[0].qty} _Carton: {ps.Where(ivn => ivn.PO == v.Split('\t')[0]).ToArray()[0].carton}\n Log-team with checking and re-share it later";
                        }
                    }
                    
                }
            }
        }
        public static void Read_PTSfile(byte[] bytes)
        {
            string title2 = "";
            using (PdfDocument document = PdfDocument.Open(bytes))
            {
                var lingw = document.GetPages().Where(page => page.GetWords().ToArray().Where(itemc => itemc.Text.Contains("Purchase")).ToArray().Length != 0);
                var array = lingw.First().GetWords().ToArray();
                int start = Array.FindIndex(array, itemc => itemc.Text.Contains("Purchase")) + 20;
                int end = Array.FindIndex(array, itemc => itemc.Text.Contains("Equipment"));
                for (int i = start + 1; i < end; i++)
                {
                    title2 = title2 + $"{array[i].Text}\t";
                    if ((i - start) % 9 == 0)
                    {
                        title2 = title2 + "\n";
                    }
                }
            }
            foreach (var vl in title2.Split('\n'))
            {
                if (vl != "")
                {
                    ps.Add(new PTSPO(vl.Split('\t')[0], vl.Split('\t')[1], vl.Split('\t')[2]));
                }
            }
        }
        public static List<PTSPO> ps = new List<PTSPO>();
        public static void Autherzire()
        {
            var tokenreq = mgss("https://network.infornexus.com/login");
            string token_name = (from item in tokenreq.Content.ReadAsStringAsync().Result.Split('<') where item.Contains("LCSRF_VAL") select item.Split('\"')[item.Split('\"').GetLength(0) - 2]).First();
            Console.WriteLine("Token string :" + token_name);
            List<KeyValuePair<string, string>> paramter = new List<KeyValuePair<string, string>>();
            paramter.Add(new KeyValuePair<string, string>("LCSRF_VAL", token_name));
            paramter.Add(new KeyValuePair<string, string>("userid", "nguyeng"));
            paramter.Add(new KeyValuePair<string, string>("userAction", ""));
            paramter.Add(new KeyValuePair<string, string>("savedMethod", ""));
            paramter.Add(new KeyValuePair<string, string>("forward", ""));
            paramter.Add(new KeyValuePair<string, string>("secureRedirectToken", ""));
            paramter.Add(new KeyValuePair<string, string>("onSuccess", ""));
            paramter.Add(new KeyValuePair<string, string>("idpHint", ""));
            paramter.Add(new KeyValuePair<string, string>("code", ""));
            paramter.Add(new KeyValuePair<string, string>("shouldRememberId", ""));
            paramter.Add(new KeyValuePair<string, string>("PAGEPERF", "1"));
            paramter.Add(new KeyValuePair<string, string>("uPassword", "A215322663b"));
            tokenreq = mgss("https://network.infornexus.com/login.jsp", paramter);
            Console.WriteLine(tokenreq.StatusCode);
        }
        public static HttpResponseMessage mgss(string url, dynamic data = null)
        {
            HttpRequestMessage req = new HttpRequestMessage(HttpMethod.Get, new Uri(url));
            if (data != null)
            {
                req = new HttpRequestMessage(HttpMethod.Post, new Uri(url));
                req.Content = new FormUrlEncodedContent(data);
            }
            var repost = client.SendAsync(req).Result;
            repost.EnsureSuccessStatusCode();
            return repost;
        }
        private static void SendEmail(string subject, string body, string recipientEmail, string cc, string[] attachmentFilePath)
        {
            Microsoft.Office.Interop.Outlook.Application outlookApp = new Microsoft.Office.Interop.Outlook.Application();
            MailItem mailItem = (MailItem)outlookApp.CreateItem(OlItemType.olMailItem);
            mailItem.Subject = subject;
            mailItem.Body = body;
            mailItem.HTMLBody = body;
            mailItem.To = recipientEmail;
            mailItem.CC = cc;
            string signature = GetSignature(outlookApp, null);
            if (!string.IsNullOrEmpty(signature))
            {
                mailItem.HTMLBody += "<br>" + signature;
            }
            foreach (var item in attachmentFilePath)
            {
                if (!string.IsNullOrEmpty(item))
                {
                    Microsoft.Office.Interop.Outlook.Attachment attachment = mailItem.Attachments.Add(item);
                }
            }
            mailItem.Send();
        }
        private static string GetSignature(Microsoft.Office.Interop.Outlook.Application outlook, string signatureName)
        {
            string signature = string.Empty;
            string sigDelimiter = "--";
            string signaturePath = string.Format("{0}\\Microsoft\\{1}\\Mail\\{2}\\{3}.html",
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "Outlook",
                outlook.Version.Substring(0, 2),
                signatureName);
            if (File.Exists(signaturePath))
            {
                signature = File.ReadAllText(signaturePath);
                signature = signature.Substring(signature.IndexOf(sigDelimiter) + sigDelimiter.Length);
                signature = signature.Replace("\n", "").Replace("\r", "");
            }
            return signature;
        }
        private static void EmailSent(ref bool Cancel)
        {
            Console.WriteLine($"Email has been sent.");
        }
        private static string[] find_file_in_path(string keyword1, string uncontains = "")
        {
            if (uncontains != "")
            {
                return Directory.GetFiles(Directory.GetCurrentDirectory(), "*", SearchOption.AllDirectories)
                                   .Where(file => File.GetLastWriteTime(file).Date == DateTime.Today.Date).Where(file => file.ToLower().Contains(keyword1.ToLower())).Where(file => !file.Contains(uncontains))
                                   .ToArray();
            }
            else
            {
                return Directory.GetFiles(Directory.GetCurrentDirectory(), "*", SearchOption.AllDirectories)
                                   .Where(file => File.GetLastWriteTime(file).Date == DateTime.Today.Date).Where(file => file.ToLower().Contains(keyword1.ToLower()))
                                   .ToArray();
            }
        }
    }
}
