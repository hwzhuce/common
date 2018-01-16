using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace Common.PDFUtil
{
    public class pdfhelp
    {
        protected void Button1_Click(object sender, EventArgs e)
        {
            PDF();
        }

        private void PDF()
        {
            string filePath = "C:\\PDF";
            if (false == Directory.Exists(filePath))
                Directory.CreateDirectory(filePath);
            string filename = filePath + "/PDF.pdf";//设置保存路径

            Document doc = new Document(iTextSharp.text.PageSize.A4, 25, 25, 50, 40);//定义pdf大小，设置上下左右边距
            PdfWriter writer = PdfWriter.GetInstance(doc, new FileStream(filename, FileMode.Create));//生成pdf路径，创建文件流

            doc.Open();
            writer.PageEvent = new HeaderAndFooterEvent();

            HeaderAndFooterEvent.PAGE_NUMBER = true;//不实现页眉跟页脚
            First(doc, writer);//封面页

            doc.NewPage();//新建一页

            PdfHeader(doc, writer);//在新建的一页里面加入数据
            HeaderAndFooterEvent.PAGE_NUMBER = false;//开始书写页眉跟页脚


            writer.Flush();
            writer.CloseStream = true;
            doc.Close();
        }

        private void PdfHeader(Document doc, PdfWriter writer)
        {
            string totalStar = string.Empty;
            writer.PageEvent = new HeaderAndFooterEvent();
            string tmp = "这个是标题";
            doc.Add(HeaderAndFooterEvent.InsertTitleContent(tmp));
        }

        private void First(Document doc, PdfWriter writer)
        {
            string tmp = "分析报告";
            doc.Add(HeaderAndFooterEvent.InsertTitleContent(tmp));

            tmp = "(正文     页,附件 0 页)";
            doc.Add(HeaderAndFooterEvent.InsertTitleContent(tmp));

            //模版 显示总共页数
            HeaderAndFooterEvent.tpl = writer.DirectContent.CreateTemplate(100, 100); //模版的宽度和高度
            PdfContentByte cb = writer.DirectContent;
            cb.AddTemplate(HeaderAndFooterEvent.tpl, 266, 714);//调节模版显示的位置

        }
    }
}
