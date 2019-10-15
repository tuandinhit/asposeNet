using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Layout;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConvertFileUtility
{
    public static class ExportFileUtility
    {
        public static byte[] AsposeTotalKey = Convert.FromBase64String(AsposeWordsLicense.AsposeTotalKey);

        public static void WordToPdf(string filePath, string savePath)
        {
            //load license key
            License license = new License();
            license.SetLicense(new MemoryStream(AsposeTotalKey));

            Document doc = new Document(filePath);
            //AddWatermark.InsertWatermarkText(doc, "Tiếng việt");
            string fileName = System.Guid.NewGuid().ToString() + ".pdf";
            doc.Save(savePath + fileName);
        }

        public static void EditFile(string filePath)
        {
            //load license key
            License license = new License();
            license.SetLicense(new MemoryStream(AsposeTotalKey));

            Document doc = new Document(Path.Combine(filePath, "Mau298NA.docx"));
            HideBookmark(doc, "DINHVANTUAN");
            doc.Save(Path.Combine(filePath, "edited.docx"));
        }

        public static void HideBookmark(Document doc, string bookMarkName)
        {
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            builder.StartBookmark(bookMarkName);
            builder.Writeln("dsada.");
            builder.EndBookmark(bookMarkName);
            builder.MoveToBookmark(bookMarkName);
        }
        public static void ShowBookmark(Document doc)
        {

        }
    }

    // For complete examples and data files, please go to https://github.com/aspose-words/Aspose.Words-for-.NET
    public static class AddWatermark
    {
        /// <summary>
        /// Inserts a watermark into a document.
        /// </summary>
        /// <param name="doc">The input document.</param>
        /// <param name="watermarkText">Text of the watermark.</param>
        public static void InsertWatermarkText(Document doc, string watermarkText)
        {
            // Create a watermark shape. This will be a WordArt shape. 
            // You are free to try other shape types as watermarks.
            Shape watermark = new Shape(doc, ShapeType.TextPlainText);
            watermark.Name = "WaterMark";
            // Set up the text of the watermark.
            watermark.TextPath.Text = watermarkText;
            watermark.TextPath.FontFamily = "Arial";
            watermark.Width = 500;
            watermark.Height = 100;
            // Text will be directed from the bottom-left to the top-right corner.
            watermark.Rotation = -40;
            // Remove the following two lines if you need a solid black text.
            watermark.Fill.Color = Color.Gray; // Try LightGray to get more Word-style watermark
            watermark.StrokeColor = Color.Gray; // Try LightGray to get more Word-style watermark

            // Place the watermark in the page center.
            watermark.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            watermark.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            watermark.WrapType = WrapType.None;
            watermark.VerticalAlignment = VerticalAlignment.Center;
            watermark.HorizontalAlignment = HorizontalAlignment.Center;

            // Create a new paragraph and append the watermark to this paragraph.
            Paragraph watermarkPara = new Paragraph(doc);
            watermarkPara.AppendChild(watermark);

            // Insert the watermark into all headers of each document section.
            foreach (Section sect in doc.Sections)
            {
                // There could be up to three different headers in each section, since we want
                // The watermark to appear on all pages, insert into all headers.
                InsertWatermarkIntoHeader(watermarkPara, sect, HeaderFooterType.HeaderPrimary);
                InsertWatermarkIntoHeader(watermarkPara, sect, HeaderFooterType.HeaderFirst);
                InsertWatermarkIntoHeader(watermarkPara, sect, HeaderFooterType.HeaderEven);
            }
        }

        private static void InsertWatermarkIntoHeader(Paragraph watermarkPara, Section sect, HeaderFooterType headerType)
        {
            HeaderFooter header = sect.HeadersFooters[headerType];

            if (header == null)
            {
                // There is no header of the specified type in the current section, create it.
                header = new HeaderFooter(sect.Document, headerType);
                sect.HeadersFooters.Add(header);
            }

            // Insert a clone of the watermark into the header.
            header.AppendChild(watermarkPara.Clone(true));
        }
    }
}
