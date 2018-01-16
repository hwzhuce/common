using System;
using System.Collections.Generic;
using System.Text;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace Common.PDFUtil
{
    public class HeaderFooterHelper : PdfPageEventHelper
    {
        public override void OnEndPage(PdfWriter writer, iTextSharp.text.Document document)
        {
              
            Rectangle rect = writer.GetBoxSize("art");
            iTextSharp.text.Image jpg = iTextSharp.text.Image.GetInstance("顺丰.jpg");    //页眉图片
            jpg.ScaleAbsolute(366, 157);   //设置图片大小
            Chunk chunk = new Chunk(jpg,-200,0);
            ColumnText.ShowTextAligned(writer.DirectContent, Element.ALIGN_RIGHT, new Phrase(chunk), rect.GetRight(10), rect.GetTop(10), 0);
            ColumnText.ShowTextAligned(writer.DirectContent, Element.ALIGN_RIGHT, new Phrase("444444444444440000000000000000000000000"), rect.GetRight(20), rect.GetTop(20), 0);
            ColumnText.ShowTextAligned(writer.DirectContent, Element.ALIGN_RIGHT, new Phrase("555555555555555550000000000000000000000000"), rect.GetRight(30), rect.GetTop(30), 0);
            //base.OnEndPage(writer, document);
        }
    }
}
