using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System.IO;
using System.Web;
using NPOI.HPSF;
using System.Data;
using NPOI.XSSF.UserModel;

namespace Common
{
    /// <summary>
    /// npoi导入excel工具类
    /// </summary>
    public static class NpoiExcelHelper
    {
        #region 数据导入到Excel
        public static IFont GetFontStyle(HSSFWorkbook workBook)
        {
            return GetFontStyle(workBook, false);
        }

        public static IFont GetFontStyle(HSSFWorkbook workBook, bool bold)
        {
            return GetFontStyle(workBook, bold, string.Empty);
        }

        public static IFont GetFontStyle(HSSFWorkbook workBook, string fontFamily)
        {
            return GetFontStyle(workBook, fontFamily, false);
        }

        public static IFont GetFontStyle(HSSFWorkbook workBook, bool bold, string fontColor)
        {
            return GetFontStyle(workBook, string.Empty, bold, fontColor);
        }

        public static IFont GetFontStyle(HSSFWorkbook workBook, string fontFamily, bool bold)
        {
            return GetFontStyle(workBook, fontFamily, string.Empty, bold, string.Empty);
        }

        public static IFont GetFontStyle(HSSFWorkbook workBook, string fontFamily, bool bold, string fontColor)
        {
            return GetFontStyle(workBook, fontFamily, string.Empty, bold, fontColor);
        }

        public static IFont GetFontStyle(HSSFWorkbook workBook, string fontFamily, string fontSize, bool bold, string fontColor)
        {
            IFont font = workBook.CreateFont();

            if (!string.IsNullOrEmpty(fontFamily))
            {
                font.FontName = fontFamily;
            }

            if (!string.IsNullOrEmpty(fontSize))
            {
                font.FontHeightInPoints = short.Parse(fontSize);
            }

            if (bold)
            {
                font.Boldweight = (short)FontBoldWeight.Bold;
            }

            if (!string.IsNullOrEmpty(fontColor))
            {
                short result;
                if (short.TryParse(fontColor, out result))
                {
                    font.Color = result;
                }
            }

            return font;
        }


        public static ICellStyle GetCellStyle(HSSFWorkbook workbook)
        {
            return GetCellStyle(workbook, HorizontalAlignment.Center, VerticalAlignment.Center);
        }

        public static ICellStyle GetCellStyle(HSSFWorkbook workbook, bool border = false, bool warpText = false)
        {
            return GetCellStyle(workbook, HorizontalAlignment.Center, VerticalAlignment.Center, border, warpText);
        }

        public static ICellStyle GetCellStyle(HSSFWorkbook workbook, HorizontalAlignment ha, VerticalAlignment va)
        {
            return GetCellStyle(workbook, ha, va, border: false, warpText: false);
        }

        public static ICellStyle GetCellStyle(HSSFWorkbook workbook, HorizontalAlignment ha, VerticalAlignment va, bool border, bool warpText)
        {
            ICellStyle style = workbook.CreateCellStyle();
            style.Alignment = ha;
            style.VerticalAlignment = va;

            if (border)
            {
                style.BorderTop = BorderStyle.Thin;
                style.BorderRight = BorderStyle.Thin;
                style.BorderBottom = BorderStyle.Thin;
                style.BorderLeft = BorderStyle.Thin;
            }

            style.WrapText = warpText;

            return style;
        }

        public static void GetMergedRegion(ISheet sheet, int rowStart, int rowEnd, int colStart, int colEnd)
        {
            CellRangeAddress region = new CellRangeAddress(rowStart, rowEnd, colStart, colEnd);

            sheet.AddMergedRegion(region);
        }

        public static string ToStr(ICell cell)
        {
            if (cell == null)
            {
                return string.Empty;
            }

            return CellTypeToValue(cell);
        }

        public static int ToInt(ICell cell)
        {
            if (cell == null)
            {
                return default(int);
            }

            string result = CellTypeToValue(cell);

            if (string.IsNullOrWhiteSpace(result))
            {
                return default(int);
            }

            return int.Parse(result);
        }

        public static DateTime? ToDateTime(ICell cell)
        {
            if (cell == null)
            {
                return null;
            }

            string result = CellTypeToValue(cell);

            if (string.IsNullOrWhiteSpace(result))
            {
                return null;
            }

            return DateTime.Parse(result);
        }

        private static string CellTypeToValue(ICell cell)
        {
            if (cell == null)
            {
                return string.Empty;
            }

            switch (cell.CellType)
            {
                case CellType.Numeric:
                    {
                        return cell.NumericCellValue.ToString();
                    }
                case CellType.Formula:
                    {
                        return cell.StringCellValue.ToString();
                    }
                case CellType.String:
                    {
                        return cell.StringCellValue.ToString();
                    }
                case CellType.Blank:
                    {
                        return string.Empty;
                    }
                default:
                    {
                        return string.Empty;
                    }
            }
        }

        /// <summary>
        /// 导出到 Excel
        /// 默认水平和垂直都是居中
        /// </summary>
        /// <param name="sheetName">sheetName</param>
        /// <param name="header">表头</param>
        /// <param name="content">内容</param>
        /// <param name="ha">水平</param>
        /// <param name="va">垂直</param>
        /// <param name="border">是否加边框</param>
        /// <param name="warpText">是否启用换行</param>
        /// <returns></returns>
        public static HSSFWorkbook ImportExcel(string sheetName, List<string> header, List<List<string>> content, HorizontalAlignment ha = HorizontalAlignment.Center, VerticalAlignment va = VerticalAlignment.Center, bool border = false, bool warpText = false)
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet(sheetName);

            ICellStyle style = GetCellStyle(workbook, ha: ha, va: va, border: border, warpText: warpText);

            int rowNumber = 0;
            IRow rowHead = sheet.CreateRow(rowNumber);
            for (int i = 0; i < header.Count; i++)
            {
                rowHead.CreateCell(i, CellType.String).SetCellValue(header[i]);
                rowHead.GetCell(i).CellStyle = style;
            }

            foreach (List<string> item in content)
            {
                ++rowNumber;
                IRow row = sheet.CreateRow(rowNumber);
                for (int i = 0; i < item.Count; i++)
                {
                    row.CreateCell(i, CellType.String).SetCellValue(item[i]);
                    row.GetCell(i).CellStyle = style;
                }
            }

            return workbook;
        }
        #endregion

        #region 导出Excel
        /// <summary>
        /// DataTable导出到Excel文件
        /// </summary>
        /// <param name="dtSource"></param>
        /// <param name="excelSavePath">保存在服务器中的物理路径</param>
        /// <param name="imgpath">如果有图片导出，图片路径</param>
        public static void ExportExcel(DataTable dtSource, string excelSavePath, string imgPath, string[] header = null)
        {
            using (MemoryStream ms = DataTableToExcel(dtSource, imgPath, header))
            {
                using (FileStream fs = new FileStream(excelSavePath, FileMode.Create, FileAccess.Write))
                {
                    byte[] data = ms.ToArray();
                    fs.Write(data, 0, data.Length);
                    fs.Flush();
                }
            }
        }

        /// <summary> 
        /// DataTable导出到Excel的MemoryStream
        /// </summary>
        /// <param name="dt">源DataTable</param>
        /// <param name="imgpath">图片路径</param>
        private static MemoryStream DataTableToExcel(DataTable dt, string imgPath, string[] header)
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFSheet sheet = (HSSFSheet)workbook.CreateSheet("sheet1");

            #region 右击文件 属性信息
            {
                //右击文件-属性-详细信息里面可见 可根据需要
                DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
                dsi.Company = "顺丰航空";
                workbook.DocumentSummaryInformation = dsi;

                SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
                si.Author = "文件作者信息"; //填加xls文件作者信息
                si.ApplicationName = "创建程序信息"; //填加xls文件创建程序信息
                si.LastAuthor = "最后保存者信息"; //填加xls文件最后保存者信息
                si.Comments = "作者信息"; //填加xls文件作者信息
                si.Title = "标题信息"; //填加xls文件标题信息
                si.Subject = "主题信息";//填加文件主题信息
                si.CreateDateTime = System.DateTime.Now;
                workbook.SummaryInformation = si;
            }
            #endregion

            HSSFCellStyle dateStyle = (HSSFCellStyle)workbook.CreateCellStyle();
            HSSFDataFormat format = (HSSFDataFormat)workbook.CreateDataFormat();
            dateStyle.DataFormat = format.GetFormat("yyyy-mm-dd");

            //取得列宽
            int[] arrColWidth = new int[dt.Columns.Count];
            foreach (DataColumn item in dt.Columns)
            {
                arrColWidth[item.Ordinal] = Encoding.GetEncoding(936).GetBytes(item.ColumnName.ToString()).Length;
            }
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    int intTemp = Encoding.GetEncoding(936).GetBytes(dt.Rows[i][j].ToString()).Length;
                    if (intTemp > arrColWidth[j])
                    {
                        arrColWidth[j] = intTemp;
                    }
                }
            }
            int rowIndex = 0;
            foreach (DataRow row in dt.Rows)
            {
                #region 新建表，填充列头，样式
                if (rowIndex == 65535 || rowIndex == 0)
                {
                    if (rowIndex != 0)
                    {
                        sheet = (HSSFSheet)workbook.CreateSheet();
                    }

                    #region 列头及样式
                    {
                        HSSFRow headerRow = (HSSFRow)sheet.CreateRow(0);
                        HSSFCellStyle headStyle = (HSSFCellStyle)workbook.CreateCellStyle();
                        //headStyle.Alignment = CellHorizontalAlignment.CENTER;
                        HSSFFont font = (HSSFFont)workbook.CreateFont();
                        font.FontHeightInPoints = 10;
                        font.Boldweight = 700;
                        headStyle.SetFont(font);

                        foreach (DataColumn column in dt.Columns)
                        {
                            //判断用户是否有自定义表头
                            if (header != null)
                            {
                                for (int i = 0; i < header.Length; i++)
                                {
                                    headerRow.CreateCell(dt.Columns[i].Ordinal).SetCellValue(header[i]);//自定义表头
                                }
                            }
                            else
                            {
                                headerRow.CreateCell(column.Ordinal).SetCellValue(column.ColumnName);//datatable自己的表头
                            }
                            headerRow.GetCell(column.Ordinal).CellStyle = headStyle;

                            //设置列宽 不设置则为自定义列宽
                            //sheet.SetColumnWidth(column.Ordinal, (arrColWidth[column.Ordinal] + 1) * 256);
                        }
                        //headerRow.Dispose();
                    }
                    #endregion

                    rowIndex = 1;
                }
                #endregion


                #region 填充内容
                HSSFRow dataRow = (HSSFRow)sheet.CreateRow(rowIndex);
                foreach (DataColumn column in dt.Columns)
                {
                    HSSFCell newCell = (HSSFCell)dataRow.CreateCell(column.Ordinal);

                    string drValue = row[column].ToString();

                    switch (column.DataType.ToString())
                    {
                        case "System.String"://字符串类型
                            newCell.SetCellValue(drValue);
                            break;
                        case "System.DateTime"://日期类型
                            System.DateTime dateV;
                            if (System.DateTime.TryParse(drValue, out dateV))
                            {
                                newCell.SetCellValue(dateV);
                                newCell.CellStyle = dateStyle;//格式化显示
                            }
                            break;
                        case "System.Boolean"://布尔型
                            bool boolV = false;
                            bool.TryParse(drValue, out boolV);
                            newCell.SetCellValue(boolV);
                            break;
                        case "System.Int16"://整型
                        case "System.Int32":
                        case "System.Int64":
                        case "System.Byte":
                            int intV = 0;
                            int.TryParse(drValue, out intV);
                            newCell.SetCellValue(intV);
                            break;
                        case "System.Decimal"://浮点型
                        case "System.Double":
                            double doubV = 0;
                            double.TryParse(drValue, out doubV);
                            newCell.SetCellValue(doubV);
                            break;
                        case "System.DBNull"://空值处理
                            newCell.SetCellValue("");
                            break;
                        default:
                            newCell.SetCellValue("");
                            break;
                    }

                }
                #endregion

                rowIndex++;
            }

            #region 导出图片
            if (imgPath != "")
            {
                HSSFSheet sheetimg = (HSSFSheet)workbook.CreateSheet("sheet2");
                imgPath = HttpContext.Current.Server.MapPath("~" + imgPath);

                byte[] bytes = System.IO.File.ReadAllBytes(imgPath);
                int pictureIdx = workbook.AddPicture(bytes, PictureType.JPEG);//PictureType.JPEG
                HSSFPatriarch patriarch = (HSSFPatriarch)sheetimg.CreateDrawingPatriarch();
                //添加图片
                //HSSFClientAnchor(dx1,dy1,dx2,dy2,col1,row1,col2,row2)的参数，有必要在这里说明一下：
                //dx1：起始单元格的x偏移量，如例子中的0表示直线起始位置距A1单元格左侧的距离；
                //dy1：起始单元格的y偏移量，如例子中的0表示直线起始位置距A1单元格上侧的距离；
                //d2：终止单元格的x偏移量，如例子中的0表示直线起始位置距C3单元格左侧的距离；
                //dy2：终止单元格的y偏移量，如例子中的150表示直线起始位置距C3单元格上侧的距离；
                //col1：起始单元格列序号，从0开始计算；
                //row1：起始单元格行序号，从0开始计算，如例子中col1=0,row1=0就表示起始单元格为A1；
                //col2：终止单元格列序号，从0开始计算；
                //row2：终止单元格行序号，从0开始计算，如例子中col2=2,row2=2就表示起始单元格为C3；
                HSSFClientAnchor anchor = new HSSFClientAnchor(0, 0, 1023, 0, 1, 4, 12, 40);
                HSSFPicture pict = (HSSFPicture)patriarch.CreatePicture(anchor, pictureIdx);
                //pict.Resize();//在结尾加上pict.Resize();即可使图片自动伸缩到原始大小了
            }
            #endregion

            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Write(ms);
                ms.Flush();
                ms.Position = 0;
                ms.Close();//关闭当前流并释放资源
                return ms;
            }
        }
        #endregion

        /// <summary>  
        /// 将excel导入到datatable  
        /// </summary>  
        /// <param name="filePath">excel路径</param>  
        /// <param name="isColumnName">第一行是否是列名</param>  
        /// <returns>返回datatable</returns>  
        public static DataTable ExcelToDataTable(string filePath, bool isColumnName)
        {
            DataTable dataTable = null;
            FileStream fs = null;
            DataColumn column = null;
            DataRow dataRow = null;
            IWorkbook workbook = null;
            ISheet sheet = null;
            IRow row = null;
            ICell cell = null;
            int startRow = 0;
            try
            {
                using (fs = File.OpenRead(filePath))
                {
                    // 2007版本  
                    if (filePath.IndexOf(".xlsx") > 0)
                        workbook = new XSSFWorkbook(fs);
                    // 2003版本  
                    else if (filePath.IndexOf(".xls") > 0)
                        workbook = new HSSFWorkbook(fs);

                    if (workbook != null)
                    {
                        IFormulaEvaluator evaluator;
                        if (filePath.IndexOf(".xlsx") > 0)
                        {
                            evaluator = new XSSFFormulaEvaluator(workbook);
                        }
                        // 2003版本  
                        else
                        {
                            evaluator = new HSSFFormulaEvaluator(workbook);
                        }
                        sheet = workbook.GetSheetAt(0);//读取第一个sheet，当然也可以循环读取每个sheet  
                        dataTable = new DataTable();
                        if (sheet != null)
                        {
                            int rowCount = sheet.LastRowNum;//总行数  
                            if (rowCount > 0)
                            {
                                IRow firstRow = sheet.GetRow(0);//第一行  
                                int cellCount = firstRow.LastCellNum;//列数  

                                //构建datatable的列  
                                if (isColumnName)
                                {
                                    startRow = 1;//如果第一行是列名，则从第二行开始读取  
                                    for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                                    {
                                        cell = firstRow.GetCell(i);
                                        if (cell != null)
                                        {
                                            if (cell.StringCellValue != null)
                                            {
                                                column = new DataColumn(cell.StringCellValue);
                                                dataTable.Columns.Add(column);
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                                    {
                                        column = new DataColumn("column" + (i + 1));
                                        dataTable.Columns.Add(column);
                                    }
                                }

                                //填充行  
                                for (int i = startRow; i <= rowCount; ++i)
                                {
                                    row = sheet.GetRow(i);
                                    if (row == null) continue;

                                    dataRow = dataTable.NewRow();
                                    for (int j = row.FirstCellNum; j < cellCount; ++j)
                                    {
                                        cell = row.GetCell(j);
                                        if (cell == null)
                                        {
                                            dataRow[j] = "";
                                        }
                                        else
                                        {
                                            //CellType(Unknown = -1,Numeric = 0,String = 1,Formula = 2,Blank = 3,Boolean = 4,Error = 5,)  
                                            switch (cell.CellType)
                                            {
                                                case CellType.Blank:
                                                    dataRow[j] = "";
                                                    break;
                                                case CellType.Numeric:
                                                    short format = cell.CellStyle.DataFormat;
                                                    //对时间格式（2015.12.5、2015/12/5、2015-12-5等）的处理  
                                                    if (format == 14 || format == 31 || format == 57 || format == 58)
                                                        dataRow[j] = cell.DateCellValue;
                                                    else
                                                        dataRow[j] = cell.NumericCellValue;
                                                    break;
                                                case CellType.String:
                                                    dataRow[j] = cell.StringCellValue;
                                                    break;
                                                case CellType.Formula:
                                                    
                                                    var cellType = evaluator.EvaluateFormulaCell(cell);
                                                    switch (cellType)
                                                    {
                                                        case CellType.Blank:
                                                            dataRow[j] = "";
                                                            break;
                                                        case CellType.Numeric:
                                                            short formulaFormat = cell.CellStyle.DataFormat;
                                                            //对时间格式（2015.12.5、2015/12/5、2015-12-5等）的处理  
                                                            if (formulaFormat == 14 || formulaFormat == 31 || formulaFormat == 57 || formulaFormat == 58)
                                                                dataRow[j] = cell.DateCellValue;
                                                            else
                                                                dataRow[j] = cell.NumericCellValue;
                                                            break;
                                                        case CellType.String:
                                                            dataRow[j] = cell.StringCellValue;
                                                            break;
                                                    }
                                                    break;
                                            }
                                        }
                                    }
                                    dataTable.Rows.Add(dataRow);
                                }
                            }
                        }
                    }
                }
                return dataTable;
            }
            catch (Exception)
            {
                if (fs != null)
                {
                    fs.Close();
                }
                return null;
            }
        }  
    }
}
