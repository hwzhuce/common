using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace CWPS.Common.PDFUtil
{
    /// <summary>
    /// PDF帮助类。by liuhy on 20151019
    /// </summary>
    public static class PDFHelper
    {

        #region 创建pdf文件
        /// <summary>
        /// 创建pdf文件。by liuhy on 20141009
        /// </summary>
        /// <param name="pdfName">包含文件路径的pdf文件名</param>
        /// <param name="pdfContent">pdf文件的内容</param>
        public static bool CreatePdfFile(string pdfName, string pdfContent, ReportPara para)
        {
            bool bResult = false;
            try
            {
                //默认字体为9号
                CreatePdfFile(pdfName, pdfContent, 9, para);
                bResult = true;
            }
            catch (Exception ex)
            {
                throw;
            }
            return bResult;
        }
        #endregion

        #region 创建pdf文件
        /// <summary>
        /// 创建pdf文件。by liuhy on 20141103
        /// </summary>
        /// <param name="pdfName">包含文件路径的pdf文件名</param>
        /// <param name="pdfContent">pdf文件的内容</param>
        /// <param name="pdfWordSize">pdf文件字体大小</param>
        public static bool CreatePdfFile(string pdfName, string pdfContent, int pdfWordSize, ReportPara para)
        {
            bool bResult = false;
            try
            {
                #region 创建pdf文件

                //获取字体
                string fontPath = @"C:\WINDOWS\Fonts\simsun.ttc";
                string mm = Environment.SystemDirectory.Substring(0, 1);
                if (!mm.ToUpper().Equals(fontPath.Substring(0, 1)))
                {
                    fontPath = mm + @":\WINDOWS\Fonts\simsun.ttc";
                }

                BaseFont baseFont = null;
                try
                {
                    baseFont = BaseFont.CreateFont(fontPath, BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
                }
                catch
                {
                    fontPath = fontPath + ",1";
                    baseFont = BaseFont.CreateFont(fontPath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                }

                Font pdfFont = new Font(baseFont, pdfWordSize, Font.NORMAL);

                Document pdfDocument = new Document(PageSize.A4, 10, 10, 100, 20);
                //PdfDestination pdfDest = new PdfDestination(PdfDestination.XYZ, 0, pdfDocument.PageSize.Height, 1f);
                PdfWriter writer = PdfWriter.GetInstance(pdfDocument, new FileStream(pdfName, FileMode.OpenOrCreate));
                //PdfAction action = PdfAction.GotoLocalPage(1, pdfDest, writer);
                if(para!=null)
                {
                    var report = new PDFReport();
                    report.ReportPar = para;
                    writer.PageEvent = report;
                }
                
                //writer.SetOpenAction(action);
                pdfDocument.Open();
                
                //HeaderAndFooterEvent.PAGE_NUMBER = true;
                
               
                
               
                //增加行距参数：10f.modify by liuhy on 20141120
                pdfDocument.Add(new Paragraph(10f, pdfContent, pdfFont));
                writer.Flush();
                writer.CloseStream = true;  
                pdfDocument.Close();
                bResult = true;

                #endregion

            }
            catch (Exception ex)
            {
                throw;
            }
            return bResult;
        }
        #endregion

        #region 通过TXT文件创建PDF文件
        /// <summary>
        /// 通过TXT文件创建PDF文件。by liuhy on 20141113
        /// </summary>
        /// <param name="pdfName">生成的pdf文件名（包含文件路径）</param>
        /// <param name="txtName">用于生成PDF的txt文件名(包含文件路径)</param>
        /// <param name="pdfWordSize">pdf文件字体大小</param>
        public static bool CreatePdfFileByTxt(string pdfName, string txtName, int pdfWordSize, ReportPara para)
        {
            bool bResult = false;
            try
            {
                string strContent = ReadFile(txtName);
                CreatePdfFile(pdfName, strContent, pdfWordSize, para);
                bResult = true;
            }
            catch (Exception ex)
            {
                throw;
            }
            return bResult;
        }
        #endregion

        #region 通过TXT文件创建PDF文件
        /// <summary>
        /// 通过TXT文件创建PDF文件。by liuhy on 20141113
        /// </summary>
        /// <param name="pdfName">生成的pdf文件名（包含文件路径）</param>
        /// <param name="txtName">用于生成PDF的txt文件名(包含文件路径)</param>
        public static bool CreatePdfFileByTxt(string pdfName, string txtName, ReportPara para)
        {
            bool bResult = false;
            try
            {
                //默认字体为9号
                CreatePdfFileByTxt(pdfName, txtName, 9, para);
                bResult = true;
            }
            catch (Exception ex)
            {
                throw;
            }
            return bResult;
        }
        #endregion

        #region 合并PDF文档
        /// <summary>
        /// 合并PDF文档。
        /// 设置文档尺寸大小。
        /// </summary>
        /// <param name="fileList">欲合并文档路径集合</param>
        /// <param name="outMergeFile">合并后的文档路径</param>
        public static bool MergePDFFiles(string[] fileList, string outMergeFile)
        {
            bool bResult = false;
            try
            {
                List<PdfReader> allFiles = new List<PdfReader>();
                PdfReader reader = null;
                Document document = new Document();
                PdfWriter writer = PdfWriter.GetInstance(document, new FileStream(outMergeFile, FileMode.OpenOrCreate));
                document.Open();
                PdfContentByte cb = writer.DirectContent;
                PdfImportedPage newPage;
                for (int i = 0; i < fileList.Length; i++)
                {
                    if (!string.IsNullOrEmpty(fileList[i]))
                    {
                        reader = new PdfReader(fileList[i]);
                        allFiles.Add(reader);
                        int iPageNum = reader.NumberOfPages;
                        for (int j = 1; j <= iPageNum; j++)
                        {
                            #region 设置文档尺寸大小.by liuhy on 20141112
                            //获取原始文档的尺寸大小
                            Rectangle newRectangle = reader.GetPageSize(j);
                            //为新的文档设置尺寸大小
                            document.SetPageSize(newRectangle);
                            #endregion

                            document.NewPage();
                            newPage = writer.GetImportedPage(reader, j);
                            cb.AddTemplate(newPage, 0, 0);
                        }
                        if (iPageNum % 2 != 0)
                        {//如果是基数页的文件，就多打印一张空白页。
                            document.SetPageSize(new Rectangle(600,850));
                            document.NewPage();
                            document.Add(new Paragraph(" "));
                            
                        }
                    }
                }
                document.Close();
                writer.Close();
                foreach (var file in allFiles)
                {
                    file.Close();
                }
                allFiles = null;
                bResult = true;
            }
            catch (Exception ex)
            {
                LogWriter.Write(LOG_CATEGORY.WEB_UI, LOG_LEVEL.ERROR, "MergePDFFiles:"+ex.Message);
                throw;
            }
            return bResult;
        }
        #endregion

        #region JPG转PDF
        /// <summary>
        /// JPG转PDF。by liuhy on 20151019
        /// </summary>
        /// <param name="pdfName">输出PDF文件路径</param>
        /// <param name="jpgName">输入JPG文件路径</param>
        public static bool CreatePdfFileByJPG(string pdfName, string jpgName,ReportPara para)
        {

            bool bResult = false;
            try
            {
                var document = new Document(iTextSharp.text.PageSize.A4, 25, 25, 100, 25);
                using (var stream = new FileStream(pdfName, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    PdfWriter writer = PdfWriter.GetInstance(document, stream);
                    if (para != null)
                    {
                        PDFReport report = new PDFReport();
                        report.ReportPar = para;
                        //report.Route = route;
                        writer.PageEvent = report;
                    }
                    document.Open();
                    using (var imageStream = new FileStream(jpgName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        var image = Image.GetInstance(imageStream);
                        if (image.Height > iTextSharp.text.PageSize.A4.Height - 25)
                        {
                            image.ScaleToFit(iTextSharp.text.PageSize.A4.Width - 25, iTextSharp.text.PageSize.A4.Height - 25);
                        }
                        else if (image.Width > iTextSharp.text.PageSize.A4.Width - 25)
                        {
                            image.ScaleToFit(iTextSharp.text.PageSize.A4.Width - 25, iTextSharp.text.PageSize.A4.Height - 25);
                        }

                        image.Alignment = iTextSharp.text.Image.ALIGN_MIDDLE;
                        document.Add(image);
                    }
                    document.Close();
                }

                bResult = true;
            }
            catch (Exception ex)
            {
            }
            return bResult;
        }
        #endregion
        

        #region CreatePdfFileByHTML
        /**********************************
        public static bool CreatePdfFileByHTML(string pdfName, string htmlName)
        {
            bool bResult = false;
            try
            {
                Document document = new Document(PageSize.A4);
                PdfWriter pdfWriter = PdfWriter.GetInstance(document, new FileStream(pdfName, FileMode.OpenOrCreate));
                pdfWriter.ViewerPreferences = PdfWriter.HideToolbar;
                //pdfWriter.InitialLeading = 12.5f;
                document.Open();
                FileStream htmlFileStream = new FileStream(htmlName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                XMLWorkerHelper helper = XMLWorkerHelper.GetInstance();
                helper.ParseXHtml(pdfWriter, document, htmlFileStream, null,Encoding.UTF8,new UnicodeFontFactory());
                document.Close();
                htmlFileStream.Close();
                bResult = true;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            return bResult;
        }
         * *********************************/

        #endregion

        #region 读文件
        /****************************************
         * 函数名称：ReadFile
         * 功能说明：读取文本内容
         * 参    数：Path:文件路径
         * 调用示列：
         *           string Path = Server.MapPath("Default2.aspx");      
         *           string s = EC.FileObj.ReadFile(Path);
        *****************************************/
        /// <summary>
        /// 读文件
        /// </summary>
        /// <param name="Path">文件路径</param>
        /// <returns></returns>
        private static string ReadFile(string Path)
        {
            string s = "";
            if (!System.IO.File.Exists(Path))
                s = "不存在相应的目录";
            else
            {
                StreamReader f2 = new StreamReader(Path, System.Text.Encoding.GetEncoding("gb2312"));
                s = f2.ReadToEnd();
                f2.Close();
                f2.Dispose();
            }

            return s;
        }
        #endregion

        /// <summary>
        /// pdf添加水印
        /// </summary>
        /// <param name="srcFile"></param>
        /// <param name="text"></param>
        /// <param name="dest"></param>
        public static void addWaterMark(String srcFile, String text,
                Stream dest)
        {
            // 待加水印的文件
            using (PdfReader reader = new PdfReader(srcFile))
            {
                //FileStream fs = new FileStream("d:/123.pdf", FileMode.OpenOrCreate);
                // 加完水印的文件
                using (PdfStamper stamper = new PdfStamper(reader, dest))
                {
                    int total = reader.NumberOfPages + 1;
                    PdfContentByte content;
                    // 设置字体
                    BaseFont font = BaseFont.CreateFont();
                    // 循环对每页插入水印
                    for (int i = 1; i < total; i++)
                    {
                        int textHeight = 60;
                        Rectangle pageSize = reader.GetPageSize(i);
                        float pageWidth = pageSize.Width;
                        float pageHeigth = pageSize.Height;

                        // 水印的起始
                        content = stamper.GetOverContent(i);
                        // 开始
                        content.BeginText();
                        while (textHeight < pageHeigth - 110)
                        {
                            int textWidth = 40;
                            while (textWidth < pageWidth - 60)
                            {
                                // 设置颜色 默认为蓝色
                                content.SetColorFill(BaseColor.GRAY);
                                // 设置字体及字号
                                content.SetFontAndSize(font, 15);
                                // 设置起始位置
                                // content.setTextMatrix(400, 880);
                                content.SaveState();
                                PdfGState gs1 = new PdfGState();
                                gs1.FillOpacity = 0.2f;
                                content.SetGState(gs1);
                                // 开始写入水印
                                content.ShowTextAligned(Element.ALIGN_LEFT, text, textWidth,
                                        textHeight, 45);
                                content.RestoreState();

                                textWidth += 60;
                            }
                            textHeight += 110;
                        }
                        content.EndText();
                    }
                    dest.Flush();
                }
            }
        }
    }
}
