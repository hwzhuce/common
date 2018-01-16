using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace Common.PDFUtil
{
    public class PDFReport : HeaderFooterHelperBase
    {
        #region 属性
        private int _pageRowCount = 12;
        /// <summary>
        /// 每页的数据行数
        /// </summary>
        public int PageRowCount
        {
            get { return _pageRowCount; }
            set { _pageRowCount = value; }
        }
        #endregion

        /// <summary>
        /// 模拟数据。实际时可能需要从数据库或网络中获取.
        /// </summary>
        

       

        #region GenerateHeader
        /// <summary>
        /// 生成页眉
        /// </summary>
        /// <param name="fontFilePath"></param>
        /// <param name="pageNumber"></param>
        /// <returns></returns>
        public override PdfPTable GenerateHeader(iTextSharp.text.pdf.PdfWriter writer)
        {
            string route = "";
            string docNameZh = "";
            string docNameEn = "";
            if(ReportPar!=null)
            {
                route = ReportPar.Route;
                docNameZh = ReportPar.DocNameZh;
                docNameEn = ReportPar.DocNameEn;
            }

            BaseFont baseFont = BaseFontForHeaderFooter;
            iTextSharp.text.Font font_logo = new iTextSharp.text.Font(baseFont, 30, iTextSharp.text.Font.BOLD);
            iTextSharp.text.Font font_header1 = new iTextSharp.text.Font(baseFont, 12, iTextSharp.text.Font.BOLD);
            iTextSharp.text.Font font_header2 = new iTextSharp.text.Font(baseFont, 12, iTextSharp.text.Font.NORMAL);
            iTextSharp.text.Font font_headerContent = new iTextSharp.text.Font(baseFont, 12, iTextSharp.text.Font.NORMAL);

            //float[] widths = new float[] { 55, 300, 50, 90, 15, 20, 15 };
            //PdfPTable header = new PdfPTable(widths);

            PdfPTable header = new PdfPTable(2);
            header.SetWidths(new float[] { 60, 300 });

            PdfPCell cell = new PdfPCell();
            //cell.BorderWidthBottom = 2;
            //cell.BorderWidthLeft = 2;
            //cell.BorderWidthTop = 2;
            //cell.BorderWidthRight = 2;
            cell.FixedHeight = 35;
            Image jpg = Image.GetInstance("顺丰.jpg");    //页眉图片
            cell = new PdfPCell(jpg, true);
            cell.HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER;
            cell.VerticalAlignment = iTextSharp.text.Element.ALIGN_CENTER;
            //cell.PaddingTop = -1;
            cell.Padding = -1;
            cell.Rowspan = 2;
            header.AddCell(cell);

            cell = GenerateOnlyBottomBorderCell(0, iTextSharp.text.Element.ALIGN_LEFT,20,8);
            cell.Phrase = new Paragraph(docNameZh, font_header1);
            header.AddCell(cell);


            cell = GenerateOnlyBottomBorderCell(0, iTextSharp.text.Element.ALIGN_LEFT,20,-2);
            cell.Phrase = new Paragraph(docNameEn, font_header1);
            header.AddCell(cell);

            cell = GenerateOnlyBottomBorderCell(0, iTextSharp.text.Element.ALIGN_LEFT, 30,0);
            cell.Phrase = new Paragraph("提取时间：", font_header2);
            header.AddCell(cell);

            cell = GenerateOnlyBottomBorderCell(0, iTextSharp.text.Element.ALIGN_LEFT,20,0);
            cell.Phrase = new Paragraph(DateTime.Now.ToString(), font_headerContent);
            header.AddCell(cell);

            cell = GenerateOnlyBottomBorderCell(2, iTextSharp.text.Element.ALIGN_LEFT,30,0);
            cell.Phrase = new Paragraph("航    线：", font_header2);
            header.AddCell(cell);

            cell = GenerateOnlyBottomBorderCell(2, iTextSharp.text.Element.ALIGN_LEFT,20,0);
            cell.Phrase = new Paragraph(route, font_headerContent);
            header.AddCell(cell);
            return header;
        }

        #region GenerateOnlyBottomBorderCell
        /// <summary>
        /// 生成只有底边的cell
        /// </summary>
        /// <param name="bottomBorder"></param>
        /// <param name="horizontalAlignment">水平对齐方式<see cref="iTextSharp.text.Element"/></param>
        /// <returns></returns>
        private PdfPCell GenerateOnlyBottomBorderCell(int bottomBorder, int horizontalAlignment, int paddingLeft, int paddingTop)
        {
            PdfPCell cell = GenerateOnlyBottomBorderCell(bottomBorder, horizontalAlignment, iTextSharp.text.Element.ALIGN_BOTTOM, paddingLeft, paddingTop);
            cell.PaddingBottom = 5;
            return cell;
        }

        /// <summary>
        /// 生成只有底边的cell
        /// </summary>
        /// <param name="bottomBorder"></param>
        /// <param name="horizontalAlignment">水平对齐方式<see cref="iTextSharp.text.Element"/></param>
        /// <param name="verticalAlignment">垂直对齐方式<see cref="iTextSharp.text.Element"/</param>
        /// <returns></returns>
        private PdfPCell GenerateOnlyBottomBorderCell(int bottomBorder, int horizontalAlignment, int verticalAlignment, int paddingLeft, int paddingTop)
        {
            PdfPCell cell = GenerateOnlyBottomBorderCell(bottomBorder, paddingLeft, paddingTop);
            cell.HorizontalAlignment = horizontalAlignment;
            cell.VerticalAlignment = verticalAlignment; ;
            return cell;
        }

        /// <summary>
        /// 生成只有底边的cell
        /// </summary>
        /// <param name="bottomBorder"></param>
        /// <param name="paddingLeft"></param>
        /// <returns></returns>
        private PdfPCell GenerateOnlyBottomBorderCell(int bottomBorder,int paddingLeft,int paddingTop)
        {
            PdfPCell cell = new PdfPCell();
            cell.BorderWidthBottom = bottomBorder;
            cell.BorderWidthLeft = 0;
            cell.BorderWidthTop = 0;
            cell.BorderWidthRight = 0;
            cell.PaddingLeft = paddingLeft;
            cell.PaddingTop = paddingTop;
            return cell;
        }
        #endregion

        #endregion

        #region GenerateFooter
        public override PdfPTable GenerateFooter(iTextSharp.text.pdf.PdfWriter writer)
        {
            BaseFont baseFont = BaseFontForHeaderFooter;
            iTextSharp.text.Font font = new iTextSharp.text.Font(baseFont, 10, iTextSharp.text.Font.NORMAL);

            PdfPTable footer = new PdfPTable(3);
            AddFooterCell(footer, "审阅:", font);
            AddFooterCell(footer, "审批:", font);
            AddFooterCell(footer, "制表:张三", font);
            return footer;
        }

        private void AddFooterCell(PdfPTable foot, String text, iTextSharp.text.Font font)
        {
            PdfPCell cell = new PdfPCell();
            cell.BorderWidthTop = 2;
            cell.BorderWidthRight = 0;
            cell.BorderWidthBottom = 0;
            cell.BorderWidthLeft = 0;
            cell.Phrase = new Phrase(text, font);
            cell.HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER;
            foot.AddCell(cell);
        }
        #endregion

        

        #region AddBody
        

       

        #region AddBodyHeader
        /// <summary>
        /// 添加Body的列标题
        /// </summary>
        /// <param name="document"></param>
        /// <param name="font_columnHeader"></param>
        private void AddBodyHeader(PdfPTable bodyTable, iTextSharp.text.Font font_columnHeader)
        {
            //采用Rowspan和Colspan来控制单元格的合并与拆分，类似于HTML的Table.
            AddColumnHeaderCell(bodyTable, "料号", font_columnHeader, 1, 2);
            AddColumnHeaderCell(bodyTable, "名称", font_columnHeader, 1, 2);
            AddColumnHeaderCell(bodyTable, "种类", font_columnHeader, 1, 2);
            AddColumnHeaderCell(bodyTable, "时间", font_columnHeader, 2, 1);//表头下还有两列表头
            AddColumnHeaderCell(bodyTable, "备注", font_columnHeader, 1, 2, true, true);

            //时间
            AddColumnHeaderCell(bodyTable, "领取", font_columnHeader);
            AddColumnHeaderCell(bodyTable, "归还", font_columnHeader, true, true);
        }
        #endregion

        

        #region AddColumnHeaderCell Function
        /// <summary>
        /// 添加列标题单元格
        /// </summary>
        /// <param name="table">表格行的单元格列表</param>
        /// <param name="header">标题</param>
        /// <param name="font">字段</param>
        /// <param name="colspan">列空间</param>
        /// <param name="rowspan">行空间</param>
        /// <param name="needLeftBorder">是否需要左边框</param>
        /// <param name="needRightBorder">是否需要右边框</param>
        public void AddColumnHeaderCell(PdfPTable table,
                                                String header,
                                                iTextSharp.text.Font font,
                                                int colspan,
                                                int rowspan,
                                                bool needLeftBorder = true,
                                                bool needRightBorder = false)
        {
            PdfPCell cell = GenerateColumnHeaderCell(header, font, needLeftBorder, needRightBorder);
            if (colspan > 1)
            {
                cell.Colspan = colspan;
            }

            if (rowspan > 1)
            {
                cell.Rowspan = rowspan;
            }

            table.AddCell(cell);
        }

        /// <summary>
        /// 添加列标题单元格
        /// </summary>
        /// <param name="table">表格</param>
        /// <param name="header">标题</param>
        /// <param name="font">字段</param>
        /// <param name="needLeftBorder">是否需要左边框</param>
        /// <param name="needRightBorder">是否需要右边框</param>
        public void AddColumnHeaderCell(PdfPTable table,
                                                String header,
                                                iTextSharp.text.Font font,
                                                bool needLeftBorder = true,
                                                bool needRightBorder = false)
        {
            PdfPCell cell = GenerateColumnHeaderCell(header, font, needLeftBorder, needRightBorder);
            table.AddCell(cell);
        }
        #endregion

        #region GenerateColumnHeaderCell
        /// <summary>
        /// 生成列标题单元格
        /// </summary>
        /// <param name="header">标题</param>
        /// <param name="font">字段</param>
        /// <param name="needLeftBorder">是否需要左边框</param>
        /// <param name="needRightBorder">是否需要右边框</param>
        /// <returns></returns>
        private PdfPCell GenerateColumnHeaderCell(String header,
                                                        iTextSharp.text.Font font,
                                                        bool needLeftBorder = true,
                                                        bool needRightBorder = false)
        {
            PdfPCell cell = new PdfPCell();
            float border = 0.5f;
            cell.BorderWidthBottom = border;
            if (needLeftBorder)
            {
                cell.BorderWidthLeft = border;
            }
            else
            {
                cell.BorderWidthLeft = 0;
            }

            cell.BorderWidthTop = border;
            if (needRightBorder)
            {
                cell.BorderWidthRight = border;
            }
            else
            {
                cell.BorderWidthRight = 0;
            }

            cell.HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER;
            cell.VerticalAlignment = iTextSharp.text.Element.ALIGN_BASELINE;
            cell.PaddingBottom = 4;
            cell.Phrase = new Phrase(header, font);
            return cell;
        }
        #endregion

        #endregion
    }
}
