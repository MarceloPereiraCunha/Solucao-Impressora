using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using iTextSharp.text.pdf;
using iTextSharp.text;
using System.IO;
using System.Runtime.InteropServices;
using System.Media;

namespace iTextSharpTest
{
    public class Etiqueta
    {
        public string Lote;
        public string Marca;
        public string DataFab;
        public string CndFab;
        public string Turno;
        public string Matricula;
        public string Peso;
        public string EstoqueAtual;
        public string Validade;
        public string Processo;
        public string TamborBag;
        public string Observacoes;
        //etiqueta.Lote, etiqueta.Marca, etiqueta.Matricula, etiqueta.Observacoes, etiqueta.Peso, processo, tambBag, etiqueta.Turno, etiqueta.CndFab, etiqueta.DataFab, etiqueta.EstoqueAtual
        public void CreatePDFDocument(string lote, string marcaMat, string matricula, string observacoesMat, string peso, string processo, string tambBag, string turno, string cndFab, string dataFab, string estoqueAtual, string validade)
        {
            //verifyFileOpened();
            
            FileStream fs = new FileStream("EtiquetaFVLS.pdf", FileMode.Create, FileAccess.Write, FileShare.None);
            Document doc = new Document();
            PdfPCell header = new PdfPCell(new Phrase("Header"));
            PdfWriter writer = PdfWriter.GetInstance(doc, fs);
            Font fontH1 = new Font();
            Font fontH2 = new Font();
            PdfPCell cell = new PdfPCell(new Phrase("Header spanning 3 columns"));
            iTextSharp.text.Image img = null;
            iTextSharp.text.Image barCodeLote = null;

            try
            {
                img = Image.GetInstance("rhiMagnesitaLogo.bmp");
                barCodeLote = Image.GetInstance("barcode.bmp");
            }

            catch (Exception ex)
            {
                throw new Exception("Falha no Carregamento de Componentes do Relatório" + ex.Message);
            }
            //fontH1.Size = 22;
            fontH1.Size = 30;
            //fontH2.Size = 10;
            fontH2.Size = 18;
            //doc.SetPageSize(iTextSharp.text.PageSize.A5.Rotate()); // for horizontal layout
            doc.SetPageSize(iTextSharp.text.PageSize.A4.Rotate()); // for horizontal layout
            
            Paragraph paragraph = new Paragraph(lote + tambBag, fontH2);
            PdfPCell cellBarcode = new PdfPCell();
            PdfPTable table = new PdfPTable(2);
            Paragraph p = new Paragraph("Matéria Prima ", fontH1);
            PdfPCell marca = new PdfPCell(new Phrase("Marca: " + marcaMat, fontH1));
            PdfPCell tamborBag = new PdfPCell(new Phrase("Tambor/Bag: ", fontH1));
            PdfPTable nested = new PdfPTable(1);
            PdfPCell nesthousing = new PdfPCell(nested);
            PdfPCell observacoes = new PdfPCell(new Phrase("Observações: " + observacoesMat, fontH1));
            

            doc.Open();

            //table.TotalWidth = 550f;
            table.TotalWidth = 680f;
            //table.TotalHeight = 900f;
            table.LockedWidth = true;

            //img.ScaleAbsolute(39f, 25f);
            img.ScaleAbsolute(59f, 38f);

            barCodeLote.ScaleAbsolute(80, 120);

            header.Colspan = 4;

            header.AddElement(img);

            p.Alignment = Element.ALIGN_MIDDLE;

            header.AddElement(p);

            table.AddCell(header);

            marca.Colspan = 14;

            tamborBag.Colspan = 14;

            table.AddCell(new PdfPCell(new Phrase("Lote: " + lote, fontH1)));

            table.AddCell(new PdfPCell(new Phrase("Tambor/BAG: " + tambBag, fontH1)));

            table.AddCell(marca);

            //nested.AddCell(new PdfPCell(new Phrase("Marca: ", fontH1)));

            nested.AddCell(new PdfPCell(new Phrase("Data Fabricação: " + dataFab, fontH1)));

            nested.AddCell(new PdfPCell(new Phrase("CND Fabricação: " + cndFab, fontH1)));

            nested.AddCell(new PdfPCell(new Phrase("Turno: " + turno, fontH1)));

            nested.AddCell(new PdfPCell(new Phrase("Matrícula: " + matricula, fontH1)));

            nested.AddCell(new PdfPCell(new Phrase("Peso: " + peso, fontH1)));

            nested.AddCell(new PdfPCell(new Phrase("Estoque Atual: " + estoqueAtual, fontH1)));

            nested.AddCell(new PdfPCell(new Phrase("Validade: " + validade, fontH1)));

            nesthousing.Padding = 0f;

            table.AddCell(nesthousing);

            observacoes.Colspan = 3;

            table.AddCell(observacoes);

            table.AddCell(new PdfPCell(new Phrase("Processo: " + processo, fontH1)));

            cellBarcode.AddElement(barCodeLote);

            cellBarcode.AddElement(paragraph);

            table.AddCell(cellBarcode);

            doc.Add(table);

            doc.Close();


        }

        //public static void verifyFileOpened()
        //{

        //    try
        //    {
        //        if (FileInUse("C:/Users/marcelo.cunha/source/repos/iTextSharpTest/iTextSharpTest/bin/Debug/Chapter1_Example1.pdf"))
        //        {
        //            File.Delete("C:/Users/marcelo.cunha/source/repos/iTextSharpTest/iTextSharpTest/bin/Debug/Chapter1_Example1.pdf");
        //        }
        //    }

        //    catch
        //    {
        //        Console.WriteLine("Feche o Arquivo");
        //    }
        //}



        //public static bool FileInUse(string path)
        //{
        //    try
        //    {
        //        using (FileStream fs = new FileStream(path, FileMode.OpenOrCreate))
        //        {
        //            return false;
        //        }
        //        //return false;
        //    }
        //    catch (IOException ex)
        //    {
        //        return true;
        //    }
    }
}
