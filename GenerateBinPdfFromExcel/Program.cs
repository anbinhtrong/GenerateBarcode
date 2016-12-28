using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using GenerateBinPdfFromExcel.Models;
using iTextSharp.text;
using iTextSharp.text.pdf;
using OfficeOpenXml;
using BarcodeLib;
using System.Drawing.Imaging;
using System.Drawing;

namespace GenerateBinPdfFromExcel
{
    class Program
    {
        static List<char> ListSpecialCharacter = new List<char>("-");
        static void Main(string[] args)
        {
            int totalfonts = FontFactory.RegisterDirectory("C:\\WINDOWS\\Fonts");

            StringBuilder sb = new StringBuilder();
            foreach (string fontname in FontFactory.RegisteredFonts)

            {

                sb.Append(fontname + "\n");

            }
            const string fileName = "Files/BinMoves(2).xlsx";
            //todo read excel file
            var bins = ReadFile(fileName);
            //todo generate pdf file from bins
            GeneratePdfFile(bins);
            //Console.ReadKey();
        }

        /// <summary>
        /// Read file from bin folder
        /// Only read the first sheet
        /// </summary>
        /// <param name="fileName"></param>
        static List<BinModel> ReadFile(string fileName)
        {
            var fileInfo = new FileInfo(fileName);
            if (fileInfo.Exists)
            {
                var package = new ExcelPackage(fileInfo);

                var workSheet = package.Workbook.Worksheets.FirstOrDefault();
                if (workSheet == null) return null;
                var row = 2;
                var bins = new List<BinModel>();
                while (row <= (workSheet.Dimension.End.Row - workSheet.Dimension.Start.Row) + 1)
                {
                    try
                    {
                        var fromBin1 = workSheet.Cells[row, 1].GetValue<string>();
                        var fromBin2 = workSheet.Cells[row, 2].GetValue<string>();
                        var toBin1 = workSheet.Cells[row, 3].GetValue<string>();
                        var toBin2 = workSheet.Cells[row, 4].GetValue<string>();
                        bins.Add(new BinModel
                        {
                            FromBin1 = fromBin1,
                            FromBin2 = fromBin2,
                            ToBin1 = toBin1,
                            ToBin2 = toBin2
                        });
                        row++;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.ToString());
                    }
                }
                return bins;
            }
            else
            {
                Console.WriteLine("File does not existed");
                return null;
            }
        }

        static void GeneratePdfFile(List<BinModel> bins)
        {
            //declare pdf file

            var baseFontTimesRoman = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, false);
            var font10 = new iTextSharp.text.Font(baseFontTimesRoman, 30f, iTextSharp.text.Font.BOLD);
            var binFont = FontFactory.GetFont("Calibri", 30f);
            var georgia = FontFactory.GetFont("Calibri", 70f);
            var verdana = FontFactory.GetFont("Verdana", 16, iTextSharp.text.Font.BOLDITALIC);
            //pagesize: 4*6
            var pageSize = new iTextSharp.text.Rectangle(432, 288);//new Rectangle(216, 72);
            var document = new Document(pageSize, 1, 1, 15, 1);
            PdfWriter.GetInstance(document, new FileStream($"BinMove2-{DateTime.Now.Ticks}.pdf", FileMode.Create));
            document.Open();
            foreach (var bin in bins)
            {

                //line 1
                var line2 = new PdfPTable(2);
                line2.DefaultCell.Padding = 0;
                line2.DefaultCell.BorderWidth = 0;
                line2.DefaultCell.Border = iTextSharp.text.Rectangle.NO_BORDER;
                line2.WidthPercentage = 100f;
                line2.PaddingTop = 72f;
                line2.SetWidths(new[] { 30, 70 });
                var cell = new PdfPCell(new Phrase("BIN #", binFont))
                {
                    HorizontalAlignment = Element.ALIGN_CENTER,
                    VerticalAlignment = Element.ALIGN_MIDDLE,
                    PaddingTop = 20,
                    BorderWidth = 0,
                    Border = iTextSharp.text.Rectangle.NO_BORDER
                };
                line2.AddCell(cell);
                var binText = bin.FromBin1;
                cell = new PdfPCell(new Phrase(binText, georgia))// &rarr;
                {
                    HorizontalAlignment = Element.ALIGN_LEFT,
                    VerticalAlignment = Element.ALIGN_MIDDLE,
                    Padding = 0,
                    BorderWidth = 0,
                    Border = iTextSharp.text.Rectangle.NO_BORDER,
                    Colspan = 2
                };
                line2.AddCell(cell);


                document.Add(line2);

                //add barcode
                var line5 = new PdfPTable(1);
                line5.WidthPercentage = 100;
                using (var barCodeStream = new MemoryStream())
                {
                    //export pdf to stream
                    GenerateBarcodeNoLable(barCodeStream, binText, imageWidth: 300, imageHeight: 60);
                    var barcodeImage = iTextSharp.text.Image.GetInstance(barCodeStream.ToArray());
                   barcodeImage.ScaleAbsoluteHeight(60f);
                    //line5.AddCell(phrase6);
                    cell = new PdfPCell(barcodeImage)// &rarr;
                    {
                        HorizontalAlignment = Element.ALIGN_CENTER,
                        VerticalAlignment = Element.ALIGN_MIDDLE,
                        PaddingTop = 80,
                        BorderWidth = 0,
                        Border = iTextSharp.text.Rectangle.NO_BORDER,
                    };
                    line5.AddCell(cell);
                }
                document.Add(line5);
                document.NewPage();
            }
            document.Close();



        }

        static void GenerateBarcodeNoLable(Stream stream, string source, TYPE type = TYPE.CODE128, int imageWidth = 200, int imageHeight = 50, string strImageFormat = "jpeg", AlignmentPositions positions = AlignmentPositions.CENTER)
        {
            try
            {
                BarcodeLib.Barcode b = new BarcodeLib.Barcode();

                if (type != TYPE.UNSPECIFIED)
                {
                    b.IncludeLabel = false;
                    b.Alignment = positions;

                    var Forecolor = "000000";
                    var Backcolor = "FFFFFF";

                    //===== Encoding performed here =====
                    System.Drawing.Image barcodeImage;
                    if (!CheckInputInSpeacial(source.Trim()))
                    {
                        barcodeImage = b.Encode(type, source.Trim(),
                           ColorTranslator.FromHtml("#" + Forecolor),
                           ColorTranslator.FromHtml("#" + Backcolor),
                           imageWidth, imageHeight);
                    }
                    else
                    {
                        barcodeImage = b.Encode(type, source.Trim(),
                       ColorTranslator.FromHtml("#" + Forecolor),
                       ColorTranslator.FromHtml("#" + Backcolor));
                    }




                    //===================================

                    //===== Static Encoding performed here =====
                    //barcodeImage = BarcodeLib.Barcode.DoEncode(type, this.txtData.Text.Trim(), this.chkGenerateLabel.Checked, this.btnForeColor.BackColor, this.btnBackColor.BackColor);
                    //==========================================

                    //Response.ContentType = "image/" + strImageFormat;

                    switch (strImageFormat)
                    {
                        case "gif": barcodeImage.Save(stream, ImageFormat.Gif); break;
                        case "jpeg": barcodeImage.Save(stream, ImageFormat.Jpeg); break;
                        case "png": barcodeImage.Save(stream, ImageFormat.Png); break;
                        case "bmp": barcodeImage.Save(stream, ImageFormat.Bmp); break;
                        case "tiff": barcodeImage.Save(stream, ImageFormat.Tiff); break;
                    }//switch

                }//if
            }//try
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }

        }
        private static bool CheckInputInSpeacial(string input)
        {
            return ListSpecialCharacter.Any(input.Contains);
        }
    }
}
