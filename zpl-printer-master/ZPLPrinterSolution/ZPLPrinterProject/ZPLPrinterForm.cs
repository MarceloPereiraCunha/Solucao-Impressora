using System;
using System.Configuration;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing.Printing;
using System.Net.Sockets;
using BarcodeLib;
using BarcodeLibTest;
using iTextSharpTest;
using iTextSharp;
using iTextSharp.text;
using SautinSoft.Document;
using PdfSharp;
using PdfSharp.Pdf;
using System.Diagnostics;
using System.Threading;


//using PdfSharp;
//using iTextSharp.text.pdf;
//using iTextSharp.text;
using Spire.Pdf;
using Spire.Pdf.Annotations;
using Spire.Pdf.Widget;


namespace ZPLPrinterProject
{
    public partial class ZPLPrinterForm : Form
    {
        private List<string> temporaryFiles = new List<string>();

        public ZPLPrinterForm(string[] args)
        {
            InitializeComponent();
            if (args.Length > 0)
            {
                OpenFile(args[0]);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // Load app settings
            var config = ConfigurationManager.OpenExeConfiguration(Application.ExecutablePath);
        }

        private void ZPLPrinterForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            // Delete temporary files
            foreach (var filePath in temporaryFiles)
            {
                File.Delete(filePath);
            }

            // Save app settings
            var config = ConfigurationManager.OpenExeConfiguration(Application.ExecutablePath);
            config.AppSettings.Settings["width"].Value = "10";
            config.AppSettings.Settings["height"].Value = "14";
            config.AppSettings.Settings["units"].Value = "cm";
            config.Save(ConfigurationSaveMode.Modified);
        }

        public void OpenFile(string path)
        {
            //sourceTextBox.Text = File.ReadAllText(path);
            //previewButton_Click(null, null);
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            //openFileDialog.Filter = "zpl files (*.zpl)|*.zpl";
            openFileDialog.FileName = "C:/Repositórios/Repositorios GIT/zpl - printer - master/ZPLPrinterSolutionzteste.zpl";


            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    OpenFile(openFileDialog.FileName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
                }
            }
        }

        private void previewButton_Click(object sender, EventArgs e)
        {
            // to inches
            double width = Double.Parse("10");
            double height = Double.Parse("14");

            width = width * 0.393701;
            height = height * 0.393701;

            byte[] zplSourceBytes = Encoding.UTF8.GetBytes(zplTemplate.ZPLCode(loteTextbox.Text, marcaTextbox.Text, tamborBagTextBox.Text, DataFabTextBox.Text, turnoTextBox.Text, matriculaTextBox.Text, pesoTextBox.Text, cndFabTextBox.Text, estoqueAtualTextBox.Text, validadeTextBox.Text, processoTextBox.Text));

            var request = (HttpWebRequest)WebRequest.Create("http://api.labelary.com/v1/printers/8dpmm/labels/" + width + "x" + height + "/");
            request.Method = "POST";
            request.Accept = "application/pdf";
            request.ContentType = "application/x-www-form-urlencoded";
            request.ContentLength = zplSourceBytes.Length;

            var requestStream = request.GetRequestStream();
            requestStream.Write(zplSourceBytes, 0, zplSourceBytes.Length);
            requestStream.Close();

            // Temp file
            string tempFileName = "temp_" + Guid.NewGuid().ToString() + ".pdf";

            temporaryFiles.Add(tempFileName);

            try
            {
                var response = (HttpWebResponse)request.GetResponse();
                var responseStream = response.GetResponseStream();
                var fileStream = File.Create(tempFileName);
                responseStream.CopyTo(fileStream);
                responseStream.Close();
                fileStream.Close();
            }
            catch (WebException exception)
            {
                Console.WriteLine("Error: {0}", exception.Status);
            }
        }

        private void printButton_Click(object sender, EventArgs e)
        {
            PrintDialog printDialog = new PrintDialog();
            printDialog.PrinterSettings = new PrinterSettings();

            string templatePrinterTag = zplTemplate.ZPLCode(loteTextbox.Text, marcaTextbox.Text, tamborBagTextBox.Text, DataFabTextBox.Text, turnoTextBox.Text, matriculaTextBox.Text, pesoTextBox.Text, cndFabTextBox.Text, estoqueAtualTextBox.Text, validadeTextBox.Text, processoTextBox.Text);
            if (templatePrinterTag != null)
            {
                if (DialogResult.OK == printDialog.ShowDialog(this))
                {
                    //If you reduce the size of the view area of the window, so the text does not all fit into one page, it will print separate pages
                    //PrintDialog printDialog = new PrintDialog();
                    switch (printDialog.PrinterSettings.PrinterName)
                    {
                        case "Foxit Reader PDF Printer":
                            // código 1
                            byte[] zpl = Encoding.UTF8.GetBytes(templatePrinterTag);

                            // adjust print density (8dpmm), label width (4 inches), label height (6 inches), and label index (0) as necessary
                            var request = (HttpWebRequest)WebRequest.Create("http://api.labelary.com/v1/printers/8dpmm/labels/4x6/0/");
                            request.Method = "POST";
                            request.Accept = "application/pdf"; // omit this line to get PNG images back
                            request.ContentType = "application/x-www-form-urlencoded";
                            request.ContentLength = zpl.Length;

                            var requestStream = request.GetRequestStream();
                            requestStream.Write(zpl, 0, zpl.Length);
                            requestStream.Close();

                            try
                            {
                                var response = (HttpWebResponse)request.GetResponse();
                                var responseStream = response.GetResponseStream();
                                var fileStream = File.Create("label.pdf"); // change file name for PNG images
                                responseStream.CopyTo(fileStream);
                                responseStream.Close();
                                fileStream.Close();
                            }
                            catch (WebException ex)
                            {
                                MessageBox.Show("Error: {0}", Convert.ToString(ex.Status));
                            }

                            break;

                        case "ZDesigner GC420t (EPL)":
                            // código 2
                            try
                            {
                                RawPrinterHelper.SendStringToPrinter(printDialog.PrinterSettings.PrinterName, templatePrinterTag);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("Falha ao enviar dado para a impressora");
                            }
                            break;

                        case "ZDesigner TLP 2844":
                            try
                            {
                                RawPrinterHelper.SendStringToPrinter(printDialog.PrinterSettings.PrinterName, templatePrinterTag);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("Falha ao enviar dado para a impressora");
                            }
                            break;

                        case "novaImpressora":
                            try
                            {
                                RawPrinterHelper.SendStringToPrinter(printDialog.PrinterSettings.PrinterName, templatePrinterTag);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("Falha ao enviar dado para a impressora");
                            }
                            break;
                    }
                }
            }
        }


        //public static void Print()
        //{
        //    PdfFilePrinter.AdobeReaderPath = @"C:\Program Files\Adobe\Reader 8.0\Reader\AcroRd32.exe";
        //    PdfFilePrinter.DefaultPrinterName = "CutePDF Writer";
        //    PdfFilePrinter printer = new PdfFilePrinter(@"C:/tmp/sample.pdf");

        //    try
        //    {
        //        printer.Print();
        //    }
        //    catch (Exception ex)
        //    {
        //        throw new NotImplementedException();
        //    }

        //}



        //private void PDFReportPrinting()
        //{
        //    //string filePath = @"C:\Solucao Impressora\zpl-printer-master\ZPLPrinterSolution\ZPLPrinterProject\bin\Release\EtiquetaFVLS.pdf";
        //    //DocumentCore doc = DocumentCore.Load(filePath);

        //    //if (doc != null)
        //    //    Console.WriteLine("Loaded successfully!");

        //    //PdfDocument doc = new PdfDocument();
        //    //doc.LoadFromFile("C:/Solucao Impressora/zpl-printer-master/ZPLPrinterSolution/ZPLPrinterProject/bin/Release/EtiquetaFVLS.pdf");

        //    byte[] fileBytes = File.ReadAllBytes(@"C:\Solucao Impressora\zpl-printer-master\ZPLPrinterSolution\ZPLPrinterProject\bin\Release\EtiquetaFVLS.pdf");

        //    DocumentCore doc = null;

        //    // Create a MemoryStream
        //    using (MemoryStream ms = new MemoryStream(fileBytes))
        //    {
        //        // Specifying PdfLoadOptions we explicitly set that a loadable document is PDF.
        //        // Also we specified here to load only 1st page and
        //        // switched off the 'OptimizeImage' to not merge adjacent images into a one.
        //        PdfLoadOptions pdfLO = new PdfLoadOptions()
        //        {
        //            PageCount = 1,
        //            OptimizeImages = false
        //        };

        //        // Load a PDF document from the MemoryStream.
        //        doc = DocumentCore.Load(ms, new PdfLoadOptions());
        //    }
        //    if (doc != null)
        //        Console.WriteLine("Loaded successfully!");

        //    //Use the default printer to print all the pages 
        //    //doc.PrintDocument.Print(); 

        //    //Set the printer and select the pages you want to print 

        //    PrintDialog dialogPrint = new PrintDialog();
        //    dialogPrint.AllowPrintToFile = true;
        //    dialogPrint.AllowSomePages = true;
        //    dialogPrint.PrinterSettings.MinimumPage = 1;
        //    //dialogPrint.PrinterSettings.MaximumPage = doc.Pages.Count;
        //    //dialogPrint.PrinterSettings.FromPage = 1;
        //    //dialogPrint.PrinterSettings.ToPage = doc.Pages.Count;
        //    doc.
        //    if (dialogPrint.ShowDialog() == DialogResult.OK)
        //    {
        //        //Set the pagenumber which you choose as the start page to print 
        //        doc.PrintFromPage = dialogPrint.PrinterSettings.FromPage;
        //        //Set the pagenumber which you choose as the final page to print 
        //        doc.PrintToPage = dialogPrint.PrinterSettings.ToPage;
        //        //Set the name of the printer which is to print the PDF 
        //        doc.PrinterName = dialogPrint.PrinterSettings.PrinterName;

        //        PrintDocument printDoc = doc.PrintDocument;
        //        dialogPrint.Document = printDoc;
        //        printDoc.Print();
        //    }
        //}


        private void PDFAdobePrinting1()
        {
            try
            {
                string Filepath = @"C:\Solucao Impressora\zpl-printer-master\ZPLPrinterSolution\ZPLPrinterProject\bin\Release\EtiquetaFVLS.pdf";

                using (PrintDialog Dialog = new PrintDialog())
                {
                    if (Dialog.ShowDialog() == DialogResult.OK)
                    {
                        //Dialog.ShowDialog();
                        ProcessStartInfo printProcessInfo = new ProcessStartInfo()
                        {
                            Verb = "print",
                            CreateNoWindow = true,
                            FileName = Filepath,
                            WindowStyle = ProcessWindowStyle.Hidden
                        };

                        Process printProcess = new Process();
                        printProcess.StartInfo = printProcessInfo;
                        printProcess.Start();

                        printProcess.WaitForInputIdle();

                        Thread.Sleep(3000);

                        if (false == printProcess.CloseMainWindow())
                        {
                            printProcess.Kill();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Envio do arquivo para a impressora falhou" + ex.Message);
            }
            

        }

        private void sourceTextBox_TextChanged(object sender, EventArgs e)
        {

        }

        public static class zplTemplate
        {
            public static string ZPLCode(string lote,
                                        string marca,
                                        string tamborBag,
                                        string dataFabricacao,
                                        string turno,
                                        string matricula,
                                        string peso,
                                        string cndDFab,
                                        string estoqueAtual,
                                        string validade,
                                        string processo)
            {
                string barCode = lote + tamborBag;
                StringBuilder etiqueta = new StringBuilder();
                etiqueta.Append("^XA");
                etiqueta.Append("^FX Top section with company logo, name and address.");
                etiqueta.Append("^CF0,60");
                etiqueta.Append("^FO50,50^GFA,4140,4140,23, 0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000070000000000000030000000000000000");
                etiqueta.Append("00000000000000FC0000000000000FC000000000000000");
                etiqueta.Append("00000000000003FF0000000000003FF000000000000000");
                etiqueta.Append("0000000000000FFF800000000000FFFC00000000000000");
                etiqueta.Append("0000000000003FFFE00000000001FFFE00000000000000");
                etiqueta.Append("000000000000FF8FF80000000007F87F80000000000000");
                etiqueta.Append("000000000001FE03FC000000001FE03FE0000000000000");
                etiqueta.Append("000000000007F800FF000000007F800FF0000000000000");
                etiqueta.Append("00000000001FE0003FC0000000FF0003FC000000000000");
                etiqueta.Append("00000000003FC0001FF0000003FC0000FF000000000000");
                etiqueta.Append("0000000000FF000007F800000FF800003FC00000000000");
                etiqueta.Append("0000000003FC000001FE00003FE000001FE00000000000");
                etiqueta.Append("000000000FF800F8007F80007F80078007F80000000000");
                etiqueta.Append("000000001FE001FC003FE001FE001FE001FE0000000000");
                etiqueta.Append("000000007F8007FF000FF007FC007FF000FF0000000000");
                etiqueta.Append("00000001FE001FFFC003FC1FF000FFFC003FC000000000");
                etiqueta.Append("00000003FC003F8FF000FF3FC003FCFF000FF000000000");
                etiqueta.Append("0000000FF000FE03F8007FFF000FF07FC003FC00000000");
                etiqueta.Append("0000001FC003FC00FE001FFC003FE01FE001FE00000000");
                etiqueta.Append("0000003F800FF0007F800FF800FF8007F8007E00000000");
                etiqueta.Append("0000001FE01FC0001FE01FE001FE0001FE001E00000000");
                etiqueta.Append("0000000FF07F000007F07F8007FC0000FF800C00000000");
                etiqueta.Append("00000003FDFE000001FDFF001FE000003FE00000000000");
                etiqueta.Append("00000000FFF8007000FFFC003FC003000FF00000000000");
                etiqueta.Append("000000007FE001FC003FF000FF000FC003FC0000000000");
                etiqueta.Append("000000001FE007FF003FE003FC003FE001FF0000000000");
                etiqueta.Append("0000000007F81FFFC0FF8007F8007FF8007F8000000000");
                etiqueta.Append("0000000003FE3F87F3FE001FE001FFFE001FE000000000");
                etiqueta.Append("0000000800FFFE03FFF8007F8007FCFF8007F800000000");
                etiqueta.Append("0000001C003FF800FFE001FF001FF03FC003FC00000000");
                etiqueta.Append("0000003F000FF0007FC003FC003FC00FF000FE00000000");
                etiqueta.Append("0000003F8003FC00FF000FF000FF0003FC003E00000000");
                etiqueta.Append("0000001FE001FE03FC003FC003FC0001FF001E00000000");
                etiqueta.Append("00000007F8007F8FF000FF800FF800007F800800000000");
                etiqueta.Append("00000001FE003FFFE001FE001FE000001FE00000000000");
                etiqueta.Append("00000000FF800FFF8007F8007F80078007F80000000000");
                etiqueta.Append("000000003FC003FE001FE001FF001FE003FC0000000000");
                etiqueta.Append("000000000FF000F8007FC003FC003FF000FF0000000000");
                etiqueta.Append("0000000007F8003000FF000FF000FFFC003FC000000000");
                etiqueta.Append("0000000001FE000003FC003FC003FFFE001FF000000000");
                etiqueta.Append("0000000C007F80000FF000FF8007F87F8007F800000000");
                etiqueta.Append("0000001E003FE0003FE001FE001FE01FE001FE00000000");
                etiqueta.Append("0000003F800FF8007F8007F8007F8007F800FE00000000");
                etiqueta.Append("0000001FC003FC01FE001FE001FF0003FE003E00000000");
                etiqueta.Append("0000000FF000FF07F8007FC003FFC00FFF000C00000000");
                etiqueta.Append("00000007FC003FDFF000FF000FF7F03FBFC00000000000");
                etiqueta.Append("00000001FE001FFFC003FC003FC1F8FE0FF00000000000");
                etiqueta.Append("000000007F8007FF000FF000FF80FFF807F80000000000");
                etiqueta.Append("000000001FE001FC001FE001FE003FF001FE0000000000");
                etiqueta.Append("000000000FF000F8007F8007FF000FC003FF8000000000");
                etiqueta.Append("0000000003FC000001FE001FFFC003000FFFE000000000");
                etiqueta.Append("0000000000FF000007F8007FCFF000003FCFF800000000");
                etiqueta.Append("0000001C003FC0001FF000FF03F80000FF03FC00000000");
                etiqueta.Append("0000003E001FF0003FC003FC00FE0001FC00FE00000000");
                etiqueta.Append("0000001F8007F800FF000FFC007F8007F800FE00000000");
                etiqueta.Append("0000001FE001FE03FE001FFE001FC01FE001FC00000000");
                etiqueta.Append("0000000FF800FF87F8007FFF8007F03F8007F800000000");
                etiqueta.Append("00000003FC003FFFE001FE3FE001FCFE001FF000000000");
                etiqueta.Append("00000000FF000FFF8007FC0FF000FFFC007FC000000000");
                etiqueta.Append("000000003FC003FF001FF003FC003FF000FF0000000000");
                etiqueta.Append("000000001FF001FC003FC000FF000FC003FC0000000000");
                etiqueta.Append("0000000007F8007000FF00007FC007800FF00000000000");
                etiqueta.Append("0000000001FE000003FC00001FE000001FE00000000000");
                etiqueta.Append("00000000007F800007F8000007F800007F800000000000");
                etiqueta.Append("00000000003FC0001FE0000001FE0001FE000000000000");
                etiqueta.Append("00000000000FF0007F80000000FF8007FC000000000000");
                etiqueta.Append("000000000003FC01FF000000003FC00FF0000000000000");
                etiqueta.Append("000000000001FF07FC000000000FF03FC0000000000000");
                etiqueta.Append("0000000000007F8FF00000000007FCFF80000000000000");
                etiqueta.Append("0000000000001FFFC00000000001FFFE00000000000000");
                etiqueta.Append("00000000000007FF8000000000007FF800000000000000");
                etiqueta.Append("00000000000003FE0000000000001FE000000000000000");
                etiqueta.Append("00000000000000F80000000000000F8000000000000000");
                etiqueta.Append("0000000000000020000000000000030000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0FE018021800400C00E001F8080307FC0FE021FFF06000");
                etiqueta.Append("1FF838061800E01C01E007FE1C030FFE1FF863FFF0E000");
                etiqueta.Append("1E7C38061800F03C01F00F8E1E070FFE3C7863FFE1F000");
                etiqueta.Append("180E380E1800F03C03F01E021F030C003008601C01B800");
                etiqueta.Append("180E380E1800F87C03303C001F870C007000601C03B800");
                etiqueta.Append("180E380E1800D86E0738380019870C003800601C031800");
                etiqueta.Append("180E380E1801DCEE0718380019C30E003C00601C031C00");
                etiqueta.Append("181C3FFE1801CCCE061C387E18E70FF81FE0601C070C00");
                etiqueta.Append("1FFC3FFE1801CFCE0E1C387E18670FF807F8601C060E00");
                etiqueta.Append("1FF038061801C78E0E1C383718770C00007C601C0E0E00");
                etiqueta.Append("183838061801C78E1FFE380618370C00001C601C0FFF00");
                etiqueta.Append("183838061801C30E1FFE3806183F0C00001C601C0FFF00");
                etiqueta.Append("181C380E1801C00E38071C06181F0C00001C601C1C0300");
                etiqueta.Append("1C1C38061801C00E38070F07180F0EC2383C601C180380");
                etiqueta.Append("180E38061801C00E300387FF18070FFE3FF8601C380380");
                etiqueta.Append("180E38061801C006700381FC18070FFE0FE0601C300380");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("0000000000000000000000000000000000000000000000");
                etiqueta.Append("^CF0,80");
                etiqueta.Append("^FO270,100^FDMateria Prima^FS");
                etiqueta.Append("^FO50,250^GB700,1,3^FS");
                etiqueta.Append("^FX Second section with recipient address and permit information.");
                etiqueta.Append("^CFA,30");
                etiqueta.Append("^FO50,300^FDMaterial: " + lote + "^FS");
                etiqueta.Append("^FO50,340^FDMarca: " + marca + "^FS");
                etiqueta.Append("^FO50,380^FDTambor/Tag: " + tamborBag + "^FS");
                etiqueta.Append("^FO50,420^FDData Fab: " + dataFabricacao + "^FS");
                etiqueta.Append("^FO50,460^FDTurno: " + turno + "^FS");
                etiqueta.Append("^FO50,500^FDMatricula: " + matricula + "^FS");
                etiqueta.Append("^FO50,540^FDPeso: " + peso + "^FS");
                etiqueta.Append("^FO50,580^FDCND Fab: " + cndDFab + "^FS");
                etiqueta.Append("^FO50,620^FDEstoque Atual: " + estoqueAtual + "^FS");
                etiqueta.Append("^FO50,660^FDValidade: " + validade + "^FS");
                etiqueta.Append("^FO50,700^FDProcesso: " + processo + "^FS");
                etiqueta.Append("^FO50,740^GB700,1,3^FS");
                etiqueta.Append("^FX Third section with barcode.");
                etiqueta.Append("^BY3,2,200");
                etiqueta.Append("^FO55,840^BC^FD" + barCode + "^FS7");
                etiqueta.Append("^FX Fourth section (the two boxes on the bottom).");
                etiqueta.Append("^XZ");
                try
                {
                    Convert.ToDouble(barCode);
                    return etiqueta.ToString();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Os campos lote e Tambor/Bag devem possuir apenas número", "Message sent to user");
                    return null;
                }
            }
        }

        private void tableLayoutPanel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void ZebraPrinting(object sender, EventArgs e)
        {
            GenerateFVLSLabel();
            PrintDialog printDialog = new PrintDialog();
            printDialog.PrinterSettings = new PrinterSettings();

            string templatePrinterTag = zplTemplate.ZPLCode(loteTextbox.Text, marcaTextbox.Text, tamborBagTextBox.Text, DataFabTextBox.Text, turnoTextBox.Text, matriculaTextBox.Text, pesoTextBox.Text, cndFabTextBox.Text, estoqueAtualTextBox.Text, validadeTextBox.Text, processoTextBox.Text);
            if (templatePrinterTag != null)
            {
                if (DialogResult.OK == printDialog.ShowDialog(this))
                {
                    switch (printDialog.PrinterSettings.PrinterName.ToString())
                    {
                        case "Foxit Reader PDF Printer":
                            // código 1
                            byte[] zpl = Encoding.UTF8.GetBytes(templatePrinterTag);

                            // adjust print density (8dpmm), label width (4 inches), label height (6 inches), and label index (0) as necessary
                            var request = (HttpWebRequest)WebRequest.Create("http://api.labelary.com/v1/printers/8dpmm/labels/4x6/0/");
                            request.Method = "POST";
                            request.Accept = "application/pdf"; // omit this line to get PNG images back
                            request.ContentType = "application/x-www-form-urlencoded";
                            request.ContentLength = zpl.Length;

                            var requestStream = request.GetRequestStream();
                            requestStream.Write(zpl, 0, zpl.Length);
                            requestStream.Close();

                            try
                            {
                                var response = (HttpWebResponse)request.GetResponse();
                                var responseStream = response.GetResponseStream();
                                var fileStream = File.Create("label.pdf"); // change file name for PNG images
                                responseStream.CopyTo(fileStream);
                                responseStream.Close();
                                fileStream.Close();
                            }
                            catch (WebException ex)
                            {
                                MessageBox.Show("Error: {0}", Convert.ToString(ex.Status));
                            }

                            break;

                        case "ZDesigner GC420t (EPL)":
                            // código 2
                            try
                            {
                                RawPrinterHelper.SendStringToPrinter(printDialog.PrinterSettings.PrinterName, templatePrinterTag);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("Falha ao enviar dado para a impressora");
                            }
                            break;

                        case "ZDesigner TLP 2844":
                            try
                            {
                                RawPrinterHelper.SendStringToPrinter(printDialog.PrinterSettings.PrinterName, templatePrinterTag);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("Falha ao enviar dado para a impressora");
                            }
                            break;

                        case "\\\\srvbhz04\\RICOH AUTOMACAO - 4510":
                            try
                            {
                                RawPrinterHelper.SendStringToPrinter(printDialog.PrinterSettings.PrinterName, templatePrinterTag);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("Falha ao enviar dado para a impressora");
                            }
                            break;


                    }
                }
            }
            if (templatePrinterTag != null)
            {
                if (DialogResult.OK == printDialog.ShowDialog(this))
                {
                    try
                    {
                        RawPrinterHelper.SendStringToPrinter(printDialog.PrinterSettings.PrinterName, templatePrinterTag);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Falha ao enviar dado para a impressora");
                    }
                }
            }
        }

        BarcodeLib.Barcode b = new BarcodeLib.Barcode();
        public void EncodeBarcode(object sender, EventArgs e)
        {
            //errorProvider1.Clear();

            int W = 300;
            int H = 150;
            b.Alignment = BarcodeLib.AlignmentPositions.CENTER;
            BarcodeLib.TYPE type = BarcodeLib.TYPE.CODE128;
            b.LabelPosition = BarcodeLib.LabelPositions.TOPCENTER;
            //barcode alignment
            try
            {
                if (type != BarcodeLib.TYPE.UNSPECIFIED)
                {
                    try
                    {
                        //b.BarWidth = textBoxBarWidth.Text.Trim().Length < 1 ? null : (int?)Convert.ToInt32(textBoxBarWidth.Text.Trim());
                        b.BarWidth = null;
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("Unable to parse BarWidth: " + ex.Message, ex);
                    }
                    try
                    {
                        //b.AspectRatio = textBoxAspectRatio.Text.Trim().Length < 1 ? null : (double?)Convert.ToDouble(textBoxAspectRatio.Text.Trim());
                        b.AspectRatio = null;
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("Unable to parse AspectRatio: " + ex.Message, ex);
                    }

                    //b.IncludeLabel = this.chkGenerateLabel.Checked;
                    b.IncludeLabel = false;

                    b.RotateFlipType = System.Drawing.RotateFlipType.Rotate180FlipXY;
                    b.LabelPosition = BarcodeLib.LabelPositions.TOPCENTER;

                    //===== Encoding performed here =====
                    //string txtDataText = "201901000000012";
                    string txtDataText = loteTextbox.Text + tamborBagTextBox.Text;
                    string ForeColorBackColor = "1";
                    //string BackColorBackColor = "0";

                    var BackgroundImage = b.Encode(type, txtDataText.Trim(), W, H);
                    //===================================

                    //show the encoding time
                    var lblEncodingTimeText = "(" + Math.Round(b.EncodingTime, 0, MidpointRounding.AwayFromZero).ToString() + "ms)";

                    var txtEncodedText = b.EncodedValue;


                }//if


            }//try
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
            }//catch
        }//btnEncode_Click

        private void SaveBarCode(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "BMP (*.bmp)|*.bmp|GIF (*.gif)|*.gif|JPG (*.jpg)|*.jpg|PNG (*.png)|*.png|TIFF (*.tif)|*.tif";
            sfd.FilterIndex = 1;
            sfd.AddExtension = true;
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                BarcodeLib.SaveTypes savetype = BarcodeLib.SaveTypes.UNSPECIFIED;
                switch (sfd.FilterIndex)
                {
                    case 1: /* BMP */  savetype = BarcodeLib.SaveTypes.BMP; break;
                    case 2: /* GIF */  savetype = BarcodeLib.SaveTypes.GIF; break;
                    case 3: /* JPG */  savetype = BarcodeLib.SaveTypes.JPG; break;
                    case 4: /* PNG */  savetype = BarcodeLib.SaveTypes.PNG; break;
                    case 5: /* TIFF */ savetype = BarcodeLib.SaveTypes.TIFF; break;
                    default: break;
                }//switch
                b.SaveImage(sfd.FileName, savetype);
            }//if
        }//btnSave_Click

        public static void verifyFileOpened()
        {
            try
            {
                if (FileInUse("C:/Users/marcelo.cunha/source/repos/iTextSharpTest/iTextSharpTest/bin/Debug/Chapter1_Example1.pdf"))
                {
                    File.Delete("C:/Users/marcelo.cunha/source/repos/iTextSharpTest/iTextSharpTest/bin/Debug/Chapter1_Example1.pdf");
                }
            }
            catch
            {
                Console.WriteLine("Feche o Arquivo");
            }
        }

        public static bool FileInUse(string path)
        {
            try
            {
                using (FileStream fs = new FileStream(path, FileMode.OpenOrCreate))
                {
                    return false;
                }
                //return false;
            }
            catch (IOException ex)
            {
                return true;
            }
        }

        private void GenerateFVLSLabel()
        {
            Etiqueta etiqueta = new Etiqueta();
            etiqueta.Lote = loteTextbox.Text;
            etiqueta.Marca = marcaTextbox.Text;
            etiqueta.Matricula = matriculaTextBox.Text;
            etiqueta.Observacoes = observacoesTextbox.Text;
            etiqueta.Peso = pesoTextBox.Text;
            etiqueta.Processo = processoTextBox.Text;
            etiqueta.TamborBag = tamborBagTextBox.Text;
            etiqueta.Turno = turnoTextBox.Text;
            etiqueta.CndFab = cndFabTextBox.Text;
            etiqueta.DataFab = DataFabTextBox.Text;
            etiqueta.EstoqueAtual = estoqueAtualTextBox.Text;
            etiqueta.Validade = validadeTextBox.Text;
            etiqueta.CreatePDFDocument(etiqueta.Lote, etiqueta.Marca, etiqueta.Matricula, etiqueta.Observacoes, etiqueta.Peso, etiqueta.Processo, etiqueta.TamborBag, etiqueta.Turno, etiqueta.CndFab, etiqueta.DataFab, etiqueta.EstoqueAtual, etiqueta.Validade);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            GenerateBarcode(sender, e);
        }

        private void GenerateBarcode(object sender, EventArgs e)
        {
            int W = 300;
            int H = 150;
            b.Alignment = BarcodeLib.AlignmentPositions.CENTER;
            //BarcodeLib.TYPE type = BarcodeLib.TYPE.CODE128;
            BarcodeLib.TYPE type = BarcodeLib.TYPE.CODE39;

            b.LabelPosition = BarcodeLib.LabelPositions.TOPCENTER;
            //barcode alignment
            try
            {
                if (type != BarcodeLib.TYPE.UNSPECIFIED)
                {
                    try
                    {
                        //b.BarWidth = textBoxBarWidth.Text.Trim().Length < 1 ? null : (int?)Convert.ToInt32(textBoxBarWidth.Text.Trim());
                        b.BarWidth = 2;

                    }
                    catch (Exception ex)
                    {
                        throw new Exception("Unable to parse BarWidth: " + ex.Message, ex);
                    }
                    try
                    {
                        //b.AspectRatio = textBoxAspectRatio.Text.Trim().Length < 1 ? null : (double?)Convert.ToDouble(textBoxAspectRatio.Text.Trim());
                        b.AspectRatio = null;
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("Unable to parse AspectRatio: " + ex.Message, ex);
                    }

                    //b.IncludeLabel = this.chkGenerateLabel.Checked;
                    b.IncludeLabel = false;

                    b.RotateFlipType = System.Drawing.RotateFlipType.Rotate180FlipXY;
                    b.LabelPosition = BarcodeLib.LabelPositions.TOPCENTER;



                    //===== Encoding performed here =====
                    string txtDataText = loteTextbox.Text + tamborBagTextBox.Text;
                    string ForeColorBackColor = "1";
                    //string BackColorBackColor = "0";

                    var BackgroundImage = b.Encode(type, txtDataText.Trim(), W, H);
                    //===================================

                    //show the encoding time
                    var lblEncodingTimeText = "(" + Math.Round(b.EncodingTime, 0, MidpointRounding.AwayFromZero).ToString() + "ms)";

                    var txtEncodedText = b.EncodedValue;

                }//if

            }//try
            catch (Exception ex)
            {
                MessageBox.Show("Os campos Lote e Tambor/BAG devem conter apenas números");
                this.Close();
                throw new Exception("Falha na geração do código de barras: " + ex.Message);
            }//catch


            //salva codigo de barras
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "BMP (*.bmp)|*.bmp|GIF (*.gif)|*.gif|JPG (*.jpg)|*.jpg|PNG (*.png)|*.png|TIFF (*.tif)|*.tif";
            sfd.FilterIndex = 1;
            sfd.AddExtension = true;
            //if (sfd.ShowDialog() == DialogResult.OK)
            //{
            BarcodeLib.SaveTypes savetype = BarcodeLib.SaveTypes.UNSPECIFIED;
            switch (sfd.FilterIndex)
            {
                case 1: /* BMP */  savetype = BarcodeLib.SaveTypes.BMP; break;
                case 2: /* GIF */  savetype = BarcodeLib.SaveTypes.GIF; break;
                case 3: /* JPG */  savetype = BarcodeLib.SaveTypes.JPG; break;
                case 4: /* PNG */  savetype = BarcodeLib.SaveTypes.PNG; break;
                case 5: /* TIFF */ savetype = BarcodeLib.SaveTypes.TIFF; break;
                default: break;
            }//switch

            sfd.FileName = "C:/Solucao Impressora/zpl-printer-master/ZPLPrinterSolution/ZPLPrinterProject/bin/Release/barcode.bmp";

            try
            {
                b.SaveImage(sfd.FileName, savetype);
            }
            catch (Exception ex)
            {
                throw new Exception("Unable to save the File: " + ex.Message, ex);
            }
            b.SaveImage(sfd.FileName, savetype);
            //}//if
        }

        private void btnGeraRelatorio_Click(object sender, EventArgs e)
        {

            if (string.IsNullOrEmpty(loteTextbox.Text) || string.IsNullOrEmpty(tamborBagTextBox.Text))
            {
                MessageBox.Show("Digite valor numérico para o Lote e Tambor/bag");
                //Application.Exit();
                this.Close();
                Application.Restart();

                //Application.Exit();
            }
            else
            {
                try
                {
                    int.Parse(loteTextbox.Text);
                    int.Parse(tamborBagTextBox.Text);
                    //Application.Restart();
                    //Application.Exit();
                }
                catch
                {
                    MessageBox.Show("Lote e Tambor/Bag deve possuir apenas caracteres numéricos");
                    this.Close();
                    Application.Restart();

                }
                finally
                {
                    //this.Close();
                    //Application.Restart();
                }
            }
            GenerateBarcode(sender, e);

            GenerateFVLSLabel();

            //PDFReportPrinting(); //comentado pq gera rótulos do SpirePDF
            PDFAdobePrinting1(); //USA ADOBE PRINT
        }

        //private bool VerifyFile()
        //{
        //    FileStream fs = null;
        //    try
        //    {
        //        fs = new FileStream("EtiquetaFVLS.pdf", FileMode.Create, FileAccess.Write, FileShare.None);

        //        return true;
        //    }
        //    catch
        //    {
        //        return false;
        //    }
        //    finally
        //    {
        //        fs.Close();
        //    }
        //}


    }
}