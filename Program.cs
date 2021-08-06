using RawPrint;
using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Text;

namespace MARKETNET_Tag_Printer
{
    class Program
    {
        const string ZINT_PATH = "C:\\Program Files (x86)\\Zint\\zint.exe";

        static int copies = 0;
        static int pages = 0;

        static void Main(string[] args)
        {
            if (args.Length == 0)
                return;

            int barcode = 0;
            string data = "";
            string[] parameters = args[0].Replace("marketnettagprinter:%5C%5C", "").Split('&');
            for (int i = 0; i < parameters.Length; i++)
            {
                string value = parameters[i].Split('=')[1];
                switch (parameters[i].Split('=')[0])
                {
                    case "copies":
                        {
                            copies = int.Parse(value);
                            break;
                        }
                    case "barcode":
                        {
                            if (value.Equals("ean13"))
                            {
                                barcode = 13;
                            }
                            else if (value.Equals("datamatrix"))
                            {
                                barcode = 71;
                            }
                            else if (value.Equals("label"))
                            {
                                barcode = -1;
                            }
                            else
                            {
                                return;
                            }
                            break;
                        }
                    case "data":
                        {
                            data = value;
                            break;
                        }
                }
            }

            if (copies == 0 || barcode == 0 || data.Length == 0)
            {
                return;
            }

            if (barcode == -1)
            {
                printLabel(data);
                return;
            }

            runZint(data, barcode);
            PrintDocument pd = new PrintDocument();
            pd.PrintPage += pd_PrintPage;
            pd.Print();
            File.Delete(Path.Combine(Path.GetTempPath(), "out.png"));
        }

        private static void runZint(string data, int barcodeType)
        {
            Process p = new Process();
            p.StartInfo.UseShellExecute = false;
            p.StartInfo.RedirectStandardOutput = true;
            p.StartInfo.FileName = ZINT_PATH;
            if (barcodeType == 71)
                p.StartInfo.Arguments = " --output=" + Path.Combine(Path.GetTempPath(), "out.png") + " -b " + barcodeType + " --square -d " + data;
            else
                p.StartInfo.Arguments = " --output=" + Path.Combine(Path.GetTempPath(), "out.png") + " -b " + barcodeType + " -d " + data;
            p.Start();
            p.WaitForExit();
        }

        private static void pd_PrintPage(object sender, PrintPageEventArgs ev)
        {
            Image img = Image.FromFile(Path.Combine(Path.GetTempPath(), "out.png"));
            ev.Graphics.DrawImage(img, 0, 0);
            img.Dispose();

            pages++;
            if (copies == pages)
                ev.HasMorePages = false;
            else
                ev.HasMorePages = true;
        }

        private static void printLabel(string data)
        {
            byte[] pdf = Convert.FromBase64String(data);
            BinaryWriter writer = new BinaryWriter(File.Open(Path.Combine(Path.GetTempPath(), "out.pdf"), FileMode.Create));
            writer.Write(pdf);
            writer.Close();

            PrinterSettings settings = new PrinterSettings();
            IPrinter printer = new Printer();
            printer.PrintRawFile(settings.PrinterName, Path.Combine(Path.GetTempPath(), "out.pdf"), "out.pdf");

            File.Delete(Path.Combine(Path.GetTempPath(), "out.pdf"));
        }

    }
}
