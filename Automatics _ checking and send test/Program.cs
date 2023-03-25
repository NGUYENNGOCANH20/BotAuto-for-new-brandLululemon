using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Interop.Excel;
using System.IO;
using static System.Net.Mime.MediaTypeNames;
using System.Runtime.InteropServices;
using static System.Net.WebRequestMethods;
using System.Diagnostics;
using File = System.IO.File;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using Page = UglyToad.PdfPig.Content.Page;
using System.Threading;
using Exception = System.Exception;

namespace Automatics___checking_and_send_test
{
    internal class Program
    {
        static void Main(string[] args)
        {
            while(true)
            {
                Task t = Task.Run(() => {
                    List<string> shipment = new List<string>();
                    shipment = Read_fileEX(find_file_in_path(".xlsx", ".xls").ToArray());
                    string was_send = File.ReadAllText(Directory.GetCurrentDirectory() + "\\Sending.txt");
                    foreach (var item in shipment)
                    {
                        if (!was_send.Contains(item))
                        {
                            Console.WriteLine($"Send test for shipment# {item}");
                            string[] attachfiles = Test_shipment(item);
                            if (attachfiles.GetLength(0)==3)
                            {
                                List<string> data = Readfile(attachfiles.Where(x => x.ToLower().Contains("packing list")).ToArray()[0]);
                                string[] dav = data.ToArray();
                                var linq_Ponumber = dav.Where(itv => Array.FindIndex(dav, v2=>v2==itv)>3 && dav[Array.FindIndex(dav, v2 => v2 == itv) - 2]=="P.O.").ToArray().Distinct();
                                var ling_Co = dav.Where(itv => Array.FindIndex(dav, v2 => v2 == itv) > 3 && dav[Array.FindIndex(dav, v2 => v2 == itv) - 2]=="CO").ToArray().Distinct();
                                string[] linepo = (from glone in dav where Array.IndexOf(dav, glone) > 2 && dav[Array.IndexOf(dav, glone) - 2] == "P.O." select glone).ToArray();
                                string[] lineco = (from glone in dav where Array.IndexOf(dav, glone) > 2 && dav[Array.IndexOf(dav, glone) - 2] == "CO" select glone).ToArray();
                                List<string> linesPCO = new List<string>();
                                List<PushPCo> poc = new List<PushPCo>();
                                for(int i = 0;i< linepo.GetLength(0); i++)
                                {
                                    linesPCO.Add(linepo[i] + "_" + lineco[i]);
                                }
                                Console.WriteLine("PO_CO :" + string.Join(" / ", linesPCO.ToArray().Distinct()));
                                int[] indexes = Enumerable.Range(0, dav.Length)
                                                .Where(i => dav[i] == "CBM:")
                                                .ToArray();
                                List<string> informations = new List<string>();
                                foreach (var item1 in indexes)
                                {
                                     double vl = 0;
                                    if(double.TryParse(dav[item1 - 8], out vl))
                                    {
                                        Console.WriteLine($"GW : {dav[item1 - 1]} _ NetW {dav[item1 - 8]} _ Pieces {dav[item1 - 12]}_ Units {dav[item1 - 15]} _ Cartons {dav[item1 - 18]} _ CBM {dav[item1 + 1]}");
                                        informations.Add($"{dav[item1 - 1]}_{dav[item1 - 8]}_{dav[item1 - 12]}_{dav[item1 - 15]}_{dav[item1 - 18]}_{dav[item1 + 1]}");
                                    }
                                    else
                                    {
                                        Console.WriteLine($"GW : {dav[item1 - 5]} _ NetW {dav[item1 - 9]} _ Pieces {dav[item1 - 13]}_ Units {dav[item1 - 16]} _ Cartons {dav[item1 - 19]} _ CBM {dav[item1 - 1]}");
                                        informations.Add($"{dav[item1-5]}_{dav[item1 - 9]}_{dav[item1 - 13]}_{dav[item1 - 16]}_{dav[item1 - 19]}_{dav[item1 - 1]}");
                                    }
                                    
                                }
                                var linePCvO = linesPCO.ToArray().Distinct().ToArray();
                                for (int i = 0; i < linePCvO.GetLength(0); i++)
                                {
                                    poc.Add(new PushPCo(linePCvO[i].Split('_')[0], linePCvO[i].Split('_')[1], informations.ToArray()[i].Split('_')[0], informations.ToArray()[i].Split('_')[1], informations.ToArray()[i].Split('_')[3], informations.ToArray()[i].Split('_')[4], informations.ToArray()[i].Split('_')[5]));
                                }
                                string subject = $"Shipment # {item} _ PO# {string.Join(" / ", linq_Ponumber)} _ CO# {string.Join(" / ", ling_Co)}";
                                string detail = "<p>INFORMATION IN PACKING LIST COMFIRMED\n</p><table border=\"1\"><tbody><tr><th>PO#</th><th>CO#</th><th>GW</th><th>NETW</th><th>Qty</th><th>Carton</th><th>CBM</th></tr>";
                                foreach (var item2 in poc)
                                {
                                    detail = detail + $"<tr><td>{item2.PO}</td><td>{item2.CO}</td><td>{item2.Gw}</td><td>{item2.Nw}</td><td>{item2.qty}</td><td>{item2.carton}</td><td>{item2.cbm}</td></tr>";
                                }
                                detail = detail + "</table>";
                                string body = $"Dear BU team ,\n<br>\n<br> Pls see and confirm Test internal INV in the attach \n<br>\n<br> Dear DGI Team ,\n<br>\n<br> Pls see and approval test original INV in system <br>\nThank you! <br>--Ai01--\n<br>\n<br>{detail}";
                                SendEmail(subject, body, File.ReadAllText(Directory.GetCurrentDirectory() + "\\to.txt"), File.ReadAllText(Directory.GetCurrentDirectory() + "\\cc.txt"), attachfiles);
                                was_send = was_send + item + "\n";
                            }
                            else
                            {
                                Console.WriteLine($"Missing test for Shipment# {item}");
                            }
                        }
                    }
                    File.WriteAllText(Directory.GetCurrentDirectory() + "\\Sending.txt", was_send);
                    Thread.Sleep(10000);
                });
                t.Wait();
            }
        }
        public class PushPCo
        {
            public string PO { get; set; }
            public string CO { get; set; }
            public string Gw { get; set; }
            public string Nw { get; set; }
            public string qty { get; set; }
            public string carton { get; set; }
            public string cbm { get; set; }
            public PushPCo(string PO, string CO,string Gw,string Nw,string qty,string carton,string cbm)
            {
                this.PO = PO;
                this.CO = CO;
                this.Gw = Gw;
                this.Nw = Nw;
                this.qty = qty;
                this.carton = carton;
                this.cbm = cbm;
            }
        }
        static List<string> Read_fileEX(string[] path)
        {
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook workbook = null;
            Worksheet worksheet = null;
            List<string> shipment = new List<string>();
            foreach (var pathc in path)
            {
                if (!pathc.Contains("$"))
                {
                    try
                    {
                        workbook = excel.Workbooks.Open(pathc);
                        foreach (var item in workbook.Sheets)
                        {
                            worksheet = (Worksheet)item;
                            Range range = worksheet.UsedRange;
                            int rowCount = range.Rows.Count;
                            int columnCount = range.Columns.Count;
                            for (int j = 1; j <= columnCount; j++)
                            {
                                if (range.Cells[1, j].Value != null && range.Cells[1, j].Value2.ToString().ToUpper().Contains("SHIPMENT"))
                                {
                                    for (int i = 2; i <= rowCount; i++)
                                    {
                                        if (range.Cells[i, j].Value != null)
                                        {
                                            shipment.Add(range.Cells[i, j].Value2.ToString());

                                        }
                                    }
                                }
                            }
                        }
                        
                    }
                    catch (System.Exception ex){ Console.WriteLine(ex.Message);}
                }
            }
            shipment.Distinct();
            workbook.Close(true);
            excel.Quit();
            GC.Collect();
            releaseObject(worksheet);
            releaseObject(workbook);
            releaseObject(excel);
            return shipment;
        }
        private static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                obj = null;
            }
            finally

            { GC.Collect(); }
        }
        private static string[] find_file_in_path(string keyword1, string keyword2, string uncontains = "")
        {
            if (uncontains != "")
            {
                //return Directory.GetFiles(File.ReadAllText($"{Directory.GetCurrentDirectory()}\\Path.txt"), "*", SearchOption.AllDirectories)
                //                   .Where(file => File.GetLastWriteTime(file).Date == DateTime.Today.Date).Where(file => file.ToLower().Contains(keyword1.ToLower())).Where(file => file.ToLower().Contains(keyword2.ToLower())).Where(file => !file.Contains(uncontains))
                //                   .ToArray();

                return Directory.GetFiles(File.ReadAllText($"{Directory.GetCurrentDirectory()}\\Path.txt"), "*", SearchOption.AllDirectories)
                                   .Where(file => file.ToLower().Contains(keyword1.ToLower())).Where(file => file.ToLower().Contains(keyword2.ToLower())).Where(file => !file.Contains(uncontains))
                                   .ToArray();
            }
            else
            {
            //    return Directory.GetFiles(File.ReadAllText($"{Directory.GetCurrentDirectory()}\\Path.txt"), "*", SearchOption.AllDirectories)
            //                       .Where(file => File.GetLastWriteTime(file).Date == DateTime.Today.Date).Where(file => file.ToLower().Contains(keyword1.ToLower())).Where(file => file.ToLower().Contains(keyword2.ToLower()))
            //                       .ToArray();
                return Directory.GetFiles(File.ReadAllText($"{Directory.GetCurrentDirectory()}\\Path.txt"), "*", SearchOption.AllDirectories)
                                   .Where(file => file.ToLower().Contains(keyword1.ToLower())).Where(file => file.ToLower().Contains(keyword2.ToLower()))
                                   .ToArray();
            }
        }
        private static string[] Test_shipment(string shipment)
        {
            return find_file_in_path(shipment, "packing list").Concat(find_file_in_path(shipment, "invoice", "000")).Distinct().ToArray();
        }
        private static void SendEmail(string subject, string body, string recipientEmail,string cc, string[] attachmentFilePath)
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
            string signaturePath = string.Format("{0}\\Microsoft\\{1}\\Mail\\{2}\\{3}.htm",
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
        private static List<string> Readfile(string pathfile)
        {
            List<string> op = new List<string>();
            using (PdfDocument document = PdfDocument.Open(pathfile))
            {
                foreach (var page in document.GetPages())
                {
                    var textContents = page.GetWords();
                    foreach (var content in textContents)
                    {
                        op.Add(content.Text);
                    }
                }
            }
            return op;
        }
    }
}
