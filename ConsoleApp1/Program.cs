using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Microsoft.Office.Interop.Outlook;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig;

namespace ConsoleApp1
{
    internal class Program
    {
        public static HttpClient client { get; set; }
        public static CookieContainer cookie { get; set; }
        public static HttpClientHandler handler { get; set; }
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
            string title = "<table border=\"1\">";
            string title2 = "<br><table border=\"1\"><tr><th>PO #</th><th>Item Quantity</th><th>Package Count</th><th>Net Weight</th><th>Net Weight</th><th>Gross</th><th>Weight Unit</th><th>Gross Volume</th><th>Volume Unit</th></tr><tr>";
            foreach (var item in PTSvalue)
            {
                string g_pts = mgss($"https://network.infornexus.com/en/trade/PlantoShipFolder?key={item}").Content.ReadAsStringAsync().Result;
                var qtitle = g_pts.Split('<').Where(INV => INV.Contains("datafieldlabelmedium"));
                if (title == "<table border=\"1\">")
                {
                    title = title + "<tr><th>PTS</th>";
                    foreach (var item1 in qtitle)
                    {
                        string vm = item1.ToString();
                        title = title + $"<th>{vm.Split('>')[1]}</th>";
                    }
                    title = title + "</tr><tr>";
                }
                var data = g_pts.Split('<').Where(INV => INV.Contains("datafieldmedium"));
                title = title + $"<td>{item}</td>";
                foreach (var item1 in data)
                {
                    string vm = item1.ToString();
                    title = title + $"<td>{vm.Split('>')[1]}</td>";
                }
                title = title + "</tr>";
                byte[] bytes = mgss($"https://network.infornexus.com/dyncon/?producer=PlatformTemplateProducer&topicName=VendorBookingRequest_viewPdf&rootId={(g_pts.Split('\"').Where(INV => INV.Contains("VendorBookingRequest?key")).ToArray()[0]).Split('=')[1]}&pmId=-1047&renderType=PDF&type=VendorBookingRequest&isHuman=true").Content.ReadAsByteArrayAsync().Result;
                File.WriteAllBytes($"{Directory.GetCurrentDirectory()}\\{item} _ PTS .pdf", bytes);
                title2 = title2 + Read_PTSfile(bytes);
                Console.WriteLine($"Done Loading PTS and save file {item}");
            }
            title = title + "</table>"+ title2+ "</table>";
            File.WriteAllText(Directory.GetCurrentDirectory() + $"\\PTS information.html", title);
            Console.WriteLine("Done export file");
            string[] attachfiles = find_file_in_path("PTS");
            SendEmail($"BOOKING PTS CREATE DATE {DateTime.Now.Date.ToString()}", $"Dear BU Team ,\n<br> Pls file PTS file in the attach <br>\nThank you! \n<br> {title}<br>--Ai02--",File.ReadAllText(Directory.GetCurrentDirectory()+"\\to.txt"),File.ReadAllText(Directory.GetCurrentDirectory() + "\\cc.txt"),attachfiles);
            Console.ReadKey();
        }
        public static string Read_PTSfile(byte[] bytes)
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
                    title2 = title2 + $"<td>{array[i].Text}</td>";
                    if ((i - start) % 9 == 0)
                    {
                        title2 = title2 + "</tr><tr>";
                    }
                }
                title2 = title2 + "</tr>";
            }
            return title2;
        }
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
