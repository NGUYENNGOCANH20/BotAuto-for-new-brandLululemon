using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using UglyToad.PdfPig;

namespace Payment__Process
{
    internal class Program
    {
        private static List<string> list = new List<string>();
        static void Main(string[] args)
        {
            Console.WriteLine("Nhap List Shipment number:");
            string[] shipments = Console.ReadLine().Split(';');
            Console.WriteLine("Nhap Bill number tuong ung:");
            string billfile = Console.ReadLine();
            FindMOM_PKL(shipments, billfile);
            int qty = Read_totalqty(list.Where(inv => inv.ToUpper().Contains("MOM")).ToArray());
            var workbook = new Aspose.Cells.Workbook(Directory.GetCurrentDirectory()+ "\\PaymentRQ.xlsx");
            var worksheet = workbook.Worksheets[0];
            worksheet.Cells[16, 0].Value = $"Shipment# {string.Join("/", shipments)}\nBill number# {billfile}\n\n\nTotal qty:{qty}";
            workbook.Save(Directory.GetCurrentDirectory() + $"\\{billfile}\\PaymentRQ_{billfile}.xlsx");
        }
        private static void FindMOM_PKL(string[] shipments, string billfile)
        {
            if(!Directory.Exists(Directory.GetCurrentDirectory() + "\\"+billfile)) {
                Directory.CreateDirectory(Directory.GetCurrentDirectory() + "\\" + billfile);
            }
            else
            {
                Console.WriteLine($"{billfile} have create folder document.");
            }
            string path = File.ReadAllText(Directory.GetCurrentDirectory() + "\\Path.txt");
            string[] momfile = Directory.GetFiles(path, "*", SearchOption.AllDirectories).Where(inv => inv.ToUpper().Contains("MOM")).ToArray();
            string[] packinglistfile = Directory.GetFiles(path, "*", SearchOption.AllDirectories).Where(inv => inv.ToUpper().Contains("PACKING")).ToArray();
            foreach (var item in shipments)
            {
                list.Add(momfile.Where(inv => inv.Contains(item)).First());
                list.Add(packinglistfile.Where(inv => inv.Contains(item)).First());
                File.Copy(momfile.Where(inv => inv.Contains(item)).First(), Directory.GetCurrentDirectory() + "\\" + billfile + "\\" + Path.GetFileName(momfile.Where(inv => inv.Contains(item)).First()), true);
                File.Copy(packinglistfile.Where(inv => inv.Contains(item)).First(), Directory.GetCurrentDirectory() + "\\" + billfile + "\\" + Path.GetFileName(packinglistfile.Where(inv => inv.Contains(item)).First()), true);
            }
            File.Copy(Directory.GetFiles(path, "*", SearchOption.AllDirectories).Where(inv => inv.Contains(billfile)).First(), Directory.GetCurrentDirectory() + "\\" + billfile + "\\" + Path.GetFileName(Directory.GetFiles(path, "*", SearchOption.AllDirectories).Where(inv => inv.Contains(billfile)).First()), true);
        }
        private static int Read_totalqty(string[] momv)
        {
            int qty = 0;
            foreach (var item in momv)
            {
                var mom = Readfile(item).ToArray();
                qty = qty + int.Parse(mom[Enumerable.Range(0, mom.Length-1).Where(i => mom[i] == "Total").ToArray()[0] + 1]);
            }
            return qty;
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
