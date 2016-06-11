//using java.util.Iterator;
//using java.util.List;
//using java.util.Map;
using org.apache.pdfbox.pdmodel;
using org.apache.pdfbox.util;
using org.apache.pdfbox.pdfparser;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using java.util;
using org.apache.pdfbox.pdmodel.graphics.xobject;
using System.Drawing;
using java.awt.image;
using java.awt;

namespace ExportPreso
{
    public class PdfEx
    {
        public static void ConvertToDoc(string pdfFilePath) {

            var pdfHeader = System.IO.Path.GetFileName(pdfFilePath);
            var underscoreIndex = pdfHeader.IndexOf('_');
            pdfHeader = pdfHeader.Remove(0, underscoreIndex + 1);

            PDDocument pdfDoc = PDDocument.load(pdfFilePath);
            var list = pdfDoc.getDocumentCatalog().getAllPages();
            Map docImages = pdfDoc.getPageMap();

            WordEx.AddTitle(pdfHeader);

            PDPage ppp = (PDPage)list.get(1);
            java.util.Iterator iter = list.iterator();
            while (iter.hasNext())
            {
                PDPage page = (PDPage)iter.next();
                
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
                        WordEx.AddImage(theImage, pdfHeader, pdfHeader);
                        theImage.Dispose();
                        //pdxObjectImage.write2file(destinationDir + fileName + "_" + totalImages);
                        //totalImages++;
                    }
                }
            }


            PDFTextStripper stripper = new PDFTextStripper();
            //stripper.setStartPage(1);
            //stripper.setEndPage(1);
           
            string docText = stripper.getText(pdfDoc);
            
            pdfDoc.close();

           

            

            WordEx.AddText(docText);

        }
    }
}
