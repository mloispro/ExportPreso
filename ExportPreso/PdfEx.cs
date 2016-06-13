//using java.util.Iterator;
//using java.util.List;
//using java.util.Map;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Drawing;
using System.IO;
using System.Drawing.Imaging;
using org.apache.pdfbox.util;
using org.apache.pdfbox.pdmodel;
using java.util;
using org.apache.pdfbox.pdmodel.graphics.xobject;
using org.apache.pdfbox.io;
using System.Windows.Forms;
using System.Text;
//using java.io;

namespace ExportPreso
{
    public class PdfEx
    {
        struct PDFPage
        {
           public string Header;
            public string Text;
        }

        static PDFTextStripper _stripper = new PDFTextStripper();
        static PDDocument _pdfDoc = null;//PDDocument.load(pdfFilePath);
        static List<Bitmap> _docImages = new List<Bitmap>();

        public static bool IsDate(string dateTime)
        {
            string[] formats = { "MM/DD/YYYY" };
            DateTime parsedDateTime;
            //bool parsedDate = DateTime.TryParse(dateTime, formats, DateTimeStyles.None, out parsedDateTime);
            //bool parsedDate = DateTime.TryParseExact(dateTime, formats, new CultureInfo("en-US"),
            //                               DateTimeStyles.None, out parsedDateTime);

            bool parsedDate = DateTime.TryParse(dateTime, out parsedDateTime);

            return parsedDate;
        }
        public static bool IsPageNumber(string number)
        {
            int num = 0;
            bool isNum = int.TryParse(number, out num);
            return isNum;
        }

        //private static List<PDFPage> ParsePages(string content)
        private static string ParsePages(string content)
        {

            ////List<PDFPage> pages = new List<PDFPage>();
            ////PDFPage page = new PDFPage();

            //if (content.Count() < 16)
            //    return content;

            //string cleanedContent = content.Remove(0, 16);
            string cleanedContent = "";
            //return cleanedContent;

            List<string> docStrings = content.Split().ToList();
            //StringBuilder pageText = new StringBuilder();

            bool headerNext = false;
            bool textNext = false;
            bool pageNumNext = false;

            foreach (string docString in docStrings)
            {
                if (docString == "")
                    continue;

                if (IsDate(docString))
                {
                    pageNumNext = true;
                    continue;
                }


                if (!IsPageNumber(docString) && pageNumNext)
                {
                    continue;
                }
                if (pageNumNext)
                {
                    pageNumNext = false;
                    continue;
                }

                //pageText.Append(docString);
                int indx = content.IndexOf(docString);
                cleanedContent = content.Substring(indx, content.Length-indx);


            }



            return cleanedContent;
        }


        ////private static List<PDFPage> ParsePages(string content)
        //private static PDFPage ParsePages(string content)
        //{
        //    //List<PDFPage> pages = new List<PDFPage>();
        //    PDFPage page = new PDFPage();
        //    List<string> docStrings = content.Split().ToList();

        //    bool headerNext = false;
        //    bool textNext = false;

        //    foreach (string docString in docStrings)
        //    {
        //        if (docString == "")
        //            continue;

        //        if (IsPageNumber(docString)) //pageBottom
        //            continue;

        //        if (IsDate(docString))
        //            continue;

        //        if (headerNext)
        //        {
        //            page.Header = docString;
        //        }
        //    }



        //    return page;
        //}

        private static List<Bitmap> GetPageImages(int pageNum)
        {
            List<Bitmap> images = new List<Bitmap>();
            return images;
        }

        private static string GetPageText(int pageNum)
        {
            _stripper.setStartPage(pageNum);
            _stripper.setEndPage(pageNum);

            string docText = _stripper.getText(_pdfDoc);
            string pageText = ParsePages(docText);

            //string pageText = pages.ToString();
            return pageText;
        }
        private static bool Equals(Bitmap bmp1, Bitmap bmp2)
        {

            byte[] image1Bytes;
            byte[] image2Bytes;

            try
            {
                using (var mstream = new MemoryStream())
                {
                    bmp1.Save(mstream, ImageFormat.Png);
                    image1Bytes = mstream.ToArray();
                }

                using (var mstream = new MemoryStream())
                {
                    bmp2.Save(mstream, ImageFormat.Png);
                    image2Bytes = mstream.ToArray();
                }
            }catch(Exception x)
            {
                return true; //return true to not add it.
            }

            var image164 = Convert.ToBase64String(image1Bytes);
            var image264 = Convert.ToBase64String(image2Bytes);

            bool equal = string.Equals(image164, image264);

            return equal;


        }
        private static bool ContainsDocImage(Bitmap bitmap)
        {
            if(_docImages.Count==0)
            {
                _docImages.Add(bitmap);
                return false;
            }
            bool shouldAdd = false;

            foreach(Bitmap img in _docImages)
            {
                if(!Equals(img, bitmap))
                {
                    shouldAdd = true;
                }else
                {
                    shouldAdd = false;//already exists.
                    break;
                }
            }

            if (shouldAdd)
            {
                _docImages.Add(bitmap);
                return false;
            }

            return true;
            
        }
        public static string ToSafeFileName(string s)
        {
            return s
                .Replace("\\", "")
                .Replace("/", "")
                .Replace("\"", "")
                .Replace("*", "")
                .Replace(":", "")
                .Replace("?", "")
                .Replace("<", "")
                .Replace(">", "")
                .Replace("|", "");
        }
        public static void ConvertToDoc(string pdfFilePath, string tempDir)
        {

            var pdfHeader = System.IO.Path.GetFileName(pdfFilePath);
            var underscoreIndex = pdfHeader.IndexOf('_');
            pdfHeader = pdfHeader.Remove(0, underscoreIndex + 1);

            WordEx.AddTitle(pdfHeader);

            _pdfDoc = new PDDocument();

            try
            {
                _pdfDoc = PDDocument.load(pdfFilePath);
            }
            catch
            {
                MessageBox.Show("Cant load pdf, try re-downloading:" + Environment.NewLine + pdfFilePath, "PDF Error");
                return;
            }

            var pagelist = _pdfDoc.getDocumentCatalog().getAllPages();

            for (int x = 0; x < pagelist.size(); x++)
            {
                //string pageTxt = GetPageText(x);
                //PDFPage pdfPage = GetPageHeader(pageTxt);

                //WordEx.AddHeader(pdfPage.Header);

                PDPage page = (PDPage)pagelist.get(x);
     
                PDResources pdResources = page.getResources();

                Map pageImages = pdResources.getImages();
                if (pageImages != null)
                {

                    Iterator imageIter = pageImages.keySet().iterator();
                    while (imageIter.hasNext())
                    {
                        String key = (String)imageIter.next();
                        PDXObjectImage pdxObjectImage = (PDXObjectImage)pageImages.get(key);
                        var buffImage = pdxObjectImage.getRGBImage();
                        Bitmap theImage = buffImage.getBitmap();
                        if (!ContainsDocImage(theImage))
                        {
                            //WordEx.AddImage(theImage, pdfPage.Header, pdfHeader);
                            WordEx.AddImage(theImage, pdfHeader, pdfHeader);
                           
                        }
                        else
                        {
                            theImage.Dispose();
                        }

                    }
                }
                
            }
            string docText = _stripper.getText(_pdfDoc);
            WordEx.AddText(docText);

            foreach (Bitmap btmap in _docImages)
            {
                btmap.Dispose();
            }
            _docImages.Clear();
            _pdfDoc.close();
            _pdfDoc = null;

        }

        private static PDFPage GetPageHeader(string pageTxt)
        {
            List<string> pageTextList = pageTxt.Split().ToList();
            string header = "";
            foreach(string txt in pageTextList)
            {
                header += txt + " ";
                if (txt == "")
                   break;
            }
            header = header.TrimEnd(" ".ToCharArray());
            string content = pageTxt.TrimStart(header.ToCharArray());
            PDFPage page = new PDFPage { Header = header, Text = content };

            return page;
        }
    }
}
