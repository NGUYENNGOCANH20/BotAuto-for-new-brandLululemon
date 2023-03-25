using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Net.Mail;
using Microsoft.Office.Interop.Outlook;
using Application = Microsoft.Office.Interop.Outlook.Application;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;

namespace Auto_Set
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            MessageBox.Show(string.Join("\n", Test_shipment("115902")));
            MessageBox.Show(string.Join("\n", DocumentM3_shipment("115902")));
        }
        private string[] fileTestINVin_path(string keyword1,string keyword2,string uncontains="")
        {
            if (uncontains != "")
            {
                return Directory.GetFiles(File.ReadAllText($"{Directory.GetCurrentDirectory()}\\Path.txt"), "*", SearchOption.AllDirectories)
                                   .Where(file => file.ToLower().Contains(keyword1.ToLower())).Where(file => file.ToLower().Contains(keyword2.ToLower())).Where(file => !file.Contains(uncontains))
                                   .ToArray();
            }
            else
            {
                return Directory.GetFiles(File.ReadAllText($"{Directory.GetCurrentDirectory()}\\Path.txt"), "*", SearchOption.AllDirectories)
                                   .Where(file => file.ToLower().Contains(keyword1.ToLower())).Where(file => file.ToLower().Contains(keyword2.ToLower()))
                                   .ToArray();
            }
        }
        private string[] Test_shipment(string shipment)
        {
            return fileTestINVin_path(shipment, "packing list").Concat(fileTestINVin_path(shipment, "invoice","000")).Distinct().ToArray();
        }
        private string[] DocumentM3_shipment(string shipment)
        {
            return fileTestINVin_path(shipment, "packing list").Concat(fileTestINVin_path(shipment, "mom")).Concat(fileTestINVin_path(shipment, "invoice")).Distinct().Except(fileTestINVin_path(shipment, "invoice", "000")).ToArray();
        }
        private void SendEmail(string subject, string body, string recipientEmail, string[] attachmentFilePath)
        {
            Application outlookApp = new Application();
            MailItem mailItem = (MailItem)outlookApp.CreateItem(OlItemType.olMailItem);
            mailItem.Subject = subject;
            mailItem.Body = body;
            mailItem.HTMLBody = body;
            mailItem.To = recipientEmail;
            string signature = GetSignature(outlookApp, null);
            if (!string.IsNullOrEmpty(signature))
            {
                mailItem.HTMLBody += "<br>" + signature;
            }
            foreach (var item in  attachmentFilePath)
            {
                if (!string.IsNullOrEmpty(item))
                {
                    Microsoft.Office.Interop.Outlook.Attachment attachment = mailItem.Attachments.Add(item);
                }
            }
            ((ItemEvents_10_Event)mailItem).Send += new ItemEvents_10_SendEventHandler(EmailSent);
            mailItem.Display(false);
        }
        private string GetSignature(Application outlook, string signatureName)
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
        private void EmailSent(ref bool Cancel)
        {
            Console.WriteLine("Email has been sent.");
        }
        private void Send_TestINV(string shipment)
        {
            SendEmail("","","", Test_shipment(shipment));
        }
        private void Send_DocumentINV(string shipment)
        {
            SendEmail("", "", "", DocumentM3_shipment(shipment));
        }
    }
}
