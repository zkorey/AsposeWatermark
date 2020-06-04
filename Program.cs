using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace AsposeWatermark
{
    class Program
    {

        /// <summary>
        /// 插入水印
        /// </summary>
        /// <param name="doc">要插入水印的文档对象</param>
        /// <param name="isTextWatermaker">是否文字水印（否的话需要传入图片的绝对地址）</param>
        /// <param name="watermarkText">水印内容（文字或图片的绝对地址）</param>
        /// <param name="rotation"></param>
        private static void InsertWatermarkText(Aspose.Words.Document doc,bool isTextWatermaker, string watermarkText,int rotation)
        {
            // Create a watermark shape. This will be a WordArt shape.
            // You are free to try other shape types as watermarks.
            Aspose.Words.Drawing.Shape watermark = null;

            if (isTextWatermaker)
            {
                watermark = new Aspose.Words.Drawing.Shape(doc, Aspose.Words.Drawing.ShapeType.TextPlainText);

                // Set up the text of the watermark.
                watermark.TextPath.Text = watermarkText;
                watermark.TextPath.FontFamily = "微软雅黑";
                watermark.Width = 500;
                watermark.Height = 100;

            }
            else {
                watermark = new Aspose.Words.Drawing.Shape(doc, Aspose.Words.Drawing.ShapeType.Image);
                watermark.ImageData.SetImage(watermarkText);
                watermark.Width =watermark.ImageData.ImageSize.WidthPixels;
                watermark.Height = watermark.ImageData.ImageSize.HeightPixels;
                watermark.HorizontalAlignment = Aspose.Words.Drawing.HorizontalAlignment.Right; //靠右对齐
                //watermark.BehindText = true;
            }
            // Text will be directed from the bottom-left to the top-right corner.
            watermark.Rotation = rotation;
            // Remove the following two lines if you need a solid black text.
            watermark.Fill.Color = System.Drawing.Color.Gray; // Try LightGray to get more Word-style watermark
            watermark.StrokeColor = System.Drawing.Color.Gray; // Try LightGray to get more Word-style watermark

            // Place the watermark in the page center.
            watermark.RelativeHorizontalPosition = Aspose.Words.Drawing.RelativeHorizontalPosition.Page;
            watermark.RelativeVerticalPosition = Aspose.Words.Drawing.RelativeVerticalPosition.Page;
            watermark.WrapType = Aspose.Words.Drawing.WrapType.None;
            watermark.VerticalAlignment = Aspose.Words.Drawing.VerticalAlignment.Center;
            watermark.HorizontalAlignment = Aspose.Words.Drawing.HorizontalAlignment.Center;

            // Create a new paragraph and append the watermark to this paragraph.
            Aspose.Words.Paragraph watermarkPara = new Aspose.Words.Paragraph(doc);
            watermarkPara.AppendChild(watermark);

            // Insert the watermark into all headers of each document section.
            foreach (Aspose.Words.Section sect in doc.Sections)
            {
                // There could be up to three different headers in each section, since we want
                // the watermark to appear on all pages, insert into all headers.
                InsertWatermarkIntoHeader(watermarkPara, sect, Aspose.Words.HeaderFooterType.HeaderPrimary);
                InsertWatermarkIntoHeader(watermarkPara, sect, Aspose.Words.HeaderFooterType.HeaderFirst);
                InsertWatermarkIntoHeader(watermarkPara, sect, Aspose.Words.HeaderFooterType.HeaderEven);
            }
        }

        private static void InsertWatermarkIntoHeader(Aspose.Words.Paragraph watermarkPara, Aspose.Words.Section sect, Aspose.Words.HeaderFooterType headerType)
        {
            Aspose.Words.HeaderFooter header = sect.HeadersFooters[headerType];

            if (header == null)
            {
                // There is no header of the specified type in the current section, create it.
                header = new Aspose.Words.HeaderFooter(sect.Document, headerType);
                sect.HeadersFooters.Add(header);
            }

            // Insert a clone of the watermark into the header.
            header.AppendChild(watermarkPara.Clone(true));
        }

        static void Main(string[] args)
        {
            string path = System.Environment.CurrentDirectory+ "\\test.docx";
            Aspose.Words.Document document2 = new Aspose.Words.Document(path);
            InsertWatermarkText(document2, false, System.Environment.CurrentDirectory+"\\123.jpg",-45);
            document2.Save(System.Environment.CurrentDirectory+"\\图片水印.3.docx");
            Console.WriteLine("添加图片水印成功");

            test2();
            Console.WriteLine("多次叠加水印成功");
            Console.ReadLine();
        }


        static void test2() {
            String WatermarkType ="text";
            String watermarkcontent ="文字水印";
            String swatermarkrotation = "-45";
            if (String.IsNullOrEmpty(WatermarkType))
            {
                WatermarkType = "text";
            }
            int watermarkrotation = 45;
            if (!String.IsNullOrEmpty(swatermarkrotation))
            {
                int.TryParse(swatermarkrotation, out watermarkrotation);
            }
            byte[] fileData = File.ReadAllBytes(System.Environment.CurrentDirectory+"\\test.docx");
            using (System.IO.MemoryStream ms = new MemoryStream(fileData))
            {
                if (!String.IsNullOrEmpty(watermarkcontent) && !String.IsNullOrEmpty(WatermarkType))
                {
                    Aspose.Words.Document document = new Aspose.Words.Document(ms);
                    InsertWatermarkText(document, !"image".Equals(WatermarkType), watermarkcontent, watermarkrotation);
                    //如果需要增加多个水印可以再次调用InsertWatermarkText 方法增加水印即可。
                    //下面这段代码是因为 跟其他接口有共用MemoryStream 对象,所以new 一个新的 MemoryStream 测试验证
                    using (System.IO.MemoryStream ms2=new MemoryStream())
                    {
                        document.Save(ms2, Aspose.Words.SaveFormat.Docx);
                        Aspose.Words.Document document2 = new Aspose.Words.Document(ms2);
                        InsertWatermarkText(document2, !"image".Equals(WatermarkType), "第二个水印", 90);
                        document2.Save(System.Environment.CurrentDirectory+"\\多次调用.4.docx");
                    }                    
                }

            }
        }
    }
}
