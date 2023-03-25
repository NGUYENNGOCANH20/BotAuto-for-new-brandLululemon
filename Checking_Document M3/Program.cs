using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using UglyToad.PdfPig;
using System.Text.RegularExpressions;
using System.Web;

namespace Checking_Document_M3
{
    internal class Program
    {
        static void Main(string[] args)
        {
            M3checking v = new M3checking();
            foreach (var item in POnumber)
            {
                Console.WriteLine("Checking PTS for PO# "+item);
                int qty = listcM3.Where(it => it.PO == item).ToArray().Sum(nv => int.Parse(nv.qty));
                int carton = listcM3.Where(it => it.PO == item).ToArray().Sum(nv => int.Parse(nv.carton));
                if (int.Parse(ps.Find(it => it.PO == item).qty) == qty)
                {
                    if (int.Parse(ps.Find(it => it.PO == item).carton) == carton)
                    {
                        Console.WriteLine($"PO# {item} _ {ps.Find(it => it.PO == item).qty} _ {ps.Find(it => it.PO == item).carton} _ Right _ Pass");
                    }
                    else
                    {
                        Console.WriteLine($"PO# {item} _ {ps.Find(it => it.PO == item).qty} Right _ Pass _ {ps.Find(it => it.PO == item).carton} _ Wrong _X");
                    }
                }
                else
                {
                    Console.WriteLine($"M3 _ PO# {item} _ {listcM3.Where(it => it.PO == item).ToArray().Sum(iv => int.Parse(iv.qty)).ToString()} EA _ {listcM3.Where(it => it.PO == item).ToArray().Sum(iv => int.Parse(iv.carton)).ToString()} Cartons");
                    Console.WriteLine($"PTS _ PO# {item} _ {ps.Find(it => it.PO == item).qty} EA _ {ps.Find(it => it.PO == item).carton} Cartons");
                }
            }
            Console.ReadKey();
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
            public PushPCo(string PO, string CO, string Gw, string Nw, string qty, string carton, string cbm)
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
        public static List<PTSPO> ps = new List<PTSPO>();
        public static List<string> POnumber = new List<string>();
        public static List<PushPCo> listcM3 = new List<PushPCo>();
        public class M3checking
        {
            public M3checking() {
                Console.WriteLine("Nhap duong dan den thu muc chua test");
                string cath = Console.ReadLine();
                string[] filePKl = find_file_in_path(cath,"packing list");
                string[] fileinternalINV = find_file_in_path(cath,"internal");
                string[] file_original = find_file_in_path(cath, "invoice").Where(item => !item.ToLower().Contains("internal")).Where(item=>!item.Contains("000")).ToArray();
                string[] filePTS = find_file_in_path(cath, "PTS");
                foreach (var item in filePTS)
                {
                    Read_PTSfile(File.ReadAllBytes(item));
                }
                foreach (var item in filePKl)
                {
                    Console.WriteLine(item);
                    List<PushPCo> list = readPKL(item);
                    foreach (var item1 in list)
                    {
                        listcM3.Add(item1);
                    }
                    string internalpath = internal_originalfile(list[0].CO,fileinternalINV);
                    string originalpath = internal_originalfile(list[0].CO, file_original);
                    int qty = list.Sum(ivn => int.Parse(ivn.qty));
                    int carton = list.Sum(ivn => int.Parse(ivn.carton));
                    string[] mom = Regex.Replace(string.Join(" ", Readfile(internalpath).ToArray()), ",", "").Split(' ');
                    if (mom.Contains(qty.ToString() + ".00"))
                    {
                        int qtyindex = Enumerable.Range(0, mom.Length).Where(i => mom[i] == qty.ToString() + ".00").ToArray()[0];
                        if (mom[qtyindex + 1] == "Transfer")
                        {
                            Console.WriteLine($"PKL #{Path.GetFileName(item)} pass qty internal test _ TEST PASS");
                        }
                        else
                        {
                            Console.WriteLine($"PKL #{Path.GetFileName(item)} wrong qty internal test");
                        }
                    }
                    else
                    {
                        Console.WriteLine($"PKL #{Path.GetFileName(item)} wrong qty internal test");
                    }
                    string[] original = Regex.Replace(string.Join(" ", Readfile(originalpath).ToArray()), ",", "").Split(' ');
                    if (original.Contains(carton.ToString()))
                    {
                        int cartonindex = Enumerable.Range(0, original.Length).Where(i => original[i] == carton.ToString()).ToArray()[0];
                        if (original[cartonindex - 3] == "Days")
                        {
                            Console.WriteLine($"PKL #{Path.GetFileName(item)} pass carton original test _ TEST PASS");
                        }
                    }
                    else
                    {
                        Console.WriteLine($"PKL #{Path.GetFileName(item)} wrong carton original test");
                        Console.WriteLine(string.Join(" ", original));
                        Console.ReadKey();
                    }
                    if (original.Contains(qty.ToString()))
                    {
                        int qtyindex = Array.FindLastIndex(original,iv=>iv== "Total:");
                        if (original[qtyindex +1] == qty.ToString())
                        {
                            Console.WriteLine($"PKL #{Path.GetFileName(item)} pass qty original test _ TEST PASS");
                        }
                        else
                        {
                            Console.WriteLine($"PKL #{Path.GetFileName(item)} wrong qty original test");
                            Console.WriteLine(string.Join(" ", original));
                            Console.ReadKey();
                        }
                    }
                    else
                    {
                        Console.WriteLine($"PKL #{Path.GetFileName(item)} wrong qty original test");
                    }
                }
            }
            public static string internal_originalfile(string check ,string[] array)
            {
                foreach(string pthg in array)
                {
                    if (Readfile(pthg).Contains(check))
                    {
                        return pthg;
                    }
                }
                return "";
            }
            public static List<PushPCo> readPKL(string packingpath)
            {
                List<PushPCo> list = new List<PushPCo>();
                string[] dav = Readfile(packingpath).ToArray();
                string[] linepo = (from glone in dav where Array.IndexOf(dav, glone) > 2 && dav[Array.IndexOf(dav, glone) - 2] == "P.O." select glone).ToArray();
                foreach (var item in linepo.Distinct())
                {
                    POnumber.Add(item);
                }
                string[] lineco = (from glone in dav where Array.IndexOf(dav, glone) > 2 && dav[Array.IndexOf(dav, glone) - 2] == "CO" select glone).ToArray();
                List<string> linesPCO = new List<string>();
                List<PushPCo> poc = new List<PushPCo>();
                for (int i = 0; i < linepo.GetLength(0); i++)
                {
                    linesPCO.Add(linepo[i] + "_" + lineco[i]);
                }
                int[] indexes = Enumerable.Range(0, dav.Length)
                                            .Where(i => dav[i] == "CBM:")
                                            .ToArray();
                List<string> informations = new List<string>();
                foreach (var item1 in indexes)
                {
                    double vl = 0;
                    if (double.TryParse(dav[item1 - 8], out vl))
                    {
                        informations.Add($"{dav[item1 - 1]}_{dav[item1 - 8]}_{dav[item1 - 12]}_{dav[item1 - 15]}_{dav[item1 - 18]}_{dav[item1 + 1]}");
                    }
                    else
                    {

                       informations.Add($"{dav[item1 - 5]}_{dav[item1 - 9]}_{dav[item1 - 13]}_{dav[item1 - 16]}_{dav[item1 - 19]}_{dav[item1-1]}");

                    }

                }
                var linePCvO = linesPCO.ToArray().Distinct().ToArray();
                for (int i = 0; i < linePCvO.GetLength(0); i++)
                {
                    list.Add(new PushPCo(linePCvO[i].Split('_')[0], linePCvO[i].Split('_')[1], informations.ToArray()[i].Split('_')[0], informations.ToArray()[i].Split('_')[1], informations.ToArray()[i].Split('_')[3], informations.ToArray()[i].Split('_')[4], informations.ToArray()[i].Split('_')[5]));
                    poc.Add(new PushPCo(linePCvO[i].Split('_')[0], linePCvO[i].Split('_')[1], informations.ToArray()[i].Split('_')[0], informations.ToArray()[i].Split('_')[1], informations.ToArray()[i].Split('_')[3], informations.ToArray()[i].Split('_')[4], informations.ToArray()[i].Split('_')[5]));
                }
                return list;
            }

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
        public static string[] find_file_in_path(string path,string keyword1, string uncontains = "")
        {
            if (uncontains != "")
            {
                return Directory.GetFiles(path, "*", SearchOption.AllDirectories)
                                   .Where(file => file.ToLower().Contains(keyword1.ToLower())).Where(file => !file.Contains(uncontains))
                                   .ToArray();
            }
            else
            {
                return Directory.GetFiles(path, "*", SearchOption.AllDirectories)
                                   .Where(file => file.ToLower().Contains(keyword1.ToLower()))
                                   .ToArray();
            }
        }
        public class PTSPO
        {
            public string PO { get; set; }
            public string qty { get; set; }
            public string carton { get; set; }
            public PTSPO(string PO,string qty, string carton)
            {
                this.PO = PO;
                this.qty = qty;
                this.carton = carton;
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
            foreach(var vl in title2.Split('\n'))
            {
                if (vl != "")
                {
                    ps.Add(new PTSPO(vl.Split('\t')[0], vl.Split('\t')[1], vl.Split('\t')[2]));
                }
                
            }
        }
    }
}
