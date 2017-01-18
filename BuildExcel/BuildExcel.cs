using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using NPOI.HSSF.Record.CF;
using NPOI.XSSF.UserModel;


namespace BuildExcel
{
    public class BuildExcel
    {

        public static int TotalColumnCoordinatePositions = 1023; //MB
        public static int TotalRowCoordinatePositions = 255; //MB
        public static int PixelsPerInch = 96; //MB
        public static double PixelsPerMillimetres = 3.78; //MB
        public static short ExcelColumnWidthFactor = 256;
        public static int UnitOffsetLength = 7;

        public static int[] UnitOffsetMap = new int[] {0, 36, 73, 109, 146, 182, 219};

        private readonly int cellPiexWidth = 72;
        private readonly int cellPiexHeight = 18;


        public BuildExcel()
        {
            this.workbook = new HSSFWorkbook();
            currentSheet = (HSSFSheet) this.workbook.CreateSheet("sheet1");
            this.workbook.CreateSheet("sheet2");
            this.workbook.CreateSheet("sheet3");

        }

        public BuildExcel(Stream fileStream)
        {
            workbook = new HSSFWorkbook(fileStream);
            this.currentSheet = (HSSFSheet) this.workbook.GetSheetAt(0);
        }

        /// <summary>
        /// NPOI文档流
        /// </summary>
        private HSSFWorkbook workbook = null;

        /// <summary>
        /// 当前操作页
        /// </summary>
        private HSSFSheet currentSheet = null;

        /// <summary>
        /// 当前单元
        /// </summary>
        private HSSFCell currentCell = null;

        /// <summary>
        /// 单元样式
        /// </summary>
        private ICellStyle cellStyle = null;



        #region--create sheet

        /// <summary>
        /// 选择操作页
        /// </summary>
        /// <param name="sheetName"></param>
        public void SelectSheet(string sheetName)
        {
            currentSheet = (HSSFSheet) workbook.GetSheet(sheetName);
        }

        /// <summary>
        /// 创建空白页
        /// </summary>
        /// <param name="sheetName"></param>
        public void CreateSheet(string sheetName)
        {
            var sheet = workbook.GetSheet(sheetName);
            if (sheet == null)
                workbook.CreateSheet(sheetName);
        }

        /// <summary>
        /// 设置页名称
        /// </summary>
        /// <param name="index"></param>
        /// <param name="sheetName"></param>
        public void SetSheetName(int index, string sheetName)
        {
            workbook.SetSheetName(index, sheetName);
        }
        #endregion

        #region--get stream

        /// <summary>
        /// 获取Excel文件流
        /// </summary>
        /// <returns></returns>
        public Stream GetStream()
        {
            MemoryStream stream = new MemoryStream();
            workbook.Write(stream);
            stream.Position = 0;
            return stream;
        }

        #endregion

        #region--sava

        public void SaveAs(string filename)
        {
            using (FileStream file = new FileStream(filename, FileMode.Create))
            {
                 workbook.Write(file);
            }
        }

        #endregion

        #region --insert text

        /// <summary>
        /// 插入文本
        /// </summary>    
        /// <param name="text">替换内容</param>
        /// <param name="bookmark">书签名称</param>
        public void InsertText(string text, string bookmark)
        {
            Replace(bookmark, text);
        }

        /// <summary>
        /// 插入文本
        /// </summary>
        /// <param name="text"></param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        public void InsertText(string text, int row, int col)
        {
            IRow r = CellUtil.GetRow(row, currentSheet);
            if (r == null)
                r = currentSheet.CreateRow(row);
            ICell cell = CellUtil.CreateCell(r, col, text, CreateStyle());
        }

        #endregion

        #region--replace text

        public void Replace(string what, string replacement)
        {
            Replace(currentSheet, what, replacement);
        }

        public void Replace(string what, string replacement, string rangeName)
        {
            var range = FindRange(rangeName);
            Replace(range, what, replacement);
        }

        private void Replace(ISheet sheet, string what, string replacement)
        {
            for (int rIndex = sheet.FirstRowNum; rIndex < sheet.LastRowNum; rIndex++)
            {
                IRow row = GetRow(rIndex);
                for (int cIndex = row.FirstCellNum; cIndex < row.LastCellNum; cIndex++)
                {
                    var cell = GetCell(row, cIndex);
                    if (cell.StringCellValue.Equals(what))
                    {
                        cell.SetCellValue(replacement);
                    }
                }
            }
        }

        private void Replace(CellRangeAddress range, string what, string replacement)
        {
            for (int rIndex = range.FirstRow; rIndex < range.LastRow; rIndex++)
            {
                IRow row = GetRow(rIndex);
                for (int cIndex = range.FirstColumn; cIndex < range.LastColumn; cIndex++)
                {
                    var cell = GetCell(row, cIndex);
                    if (cell.StringCellValue.Equals(what))
                    {
                        cell.SetCellValue(replacement);
                    }
                }
            }
        }

        #endregion

        #region--GetBookmarks，GetAllMarks，DelBookmarks

        /// <summary>
        /// 获取所有书签
        /// </summary>
        /// <returns></returns>
        public List<string> GetBookmarks()
        {
            return FindAllText(currentSheet, @"《([^》]+)》");
        }

        public List<string> GetBookmarks(string rangeName)
        {
            var range = FindRange(rangeName);
            return FindAllText(range, @"《([^》]+)》");
        }

        #endregion

        #region --insert table

        /// <summary>
        /// 插入表格
        /// </summary>
        /// <param name="table"></param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        public void InsertTable(DataTable table, int row, int col)
        {
            ICell cell = GetCell(row, col);
            int rowIndex = cell.RowIndex;
            InsertTable(table, rowIndex, true);

        }


        public void ReplaceInsertTable(string what, DataTable table)
        {
            ICell cell = Find(currentSheet, what);
            if (cell != null)
            {
                int rowCount = table.Rows.Count;
                int rowIndex = cell.RowIndex;
                ShiftRows(currentSheet, rowIndex + 1, rowIndex + rowCount, rowCount - 1); //keep what row
                InsertTable(table, rowIndex, false);
            }
        }

        public void InsertTable(DataTable table, int rowIndex, bool hasHeader)
        {
            InsertTable(table, rowIndex, hasHeader, GetThinBDRStyle());
        }

        private void InsertTable(DataTable table, int rowIndex, bool hasHeader, ICellStyle style)
        {
            
            var sheet = currentSheet;
            if (hasHeader)
            {
                var headerRow = sheet.CreateRow(rowIndex);
                foreach (DataColumn column in table.Columns)
                {
                    var headerCell = headerRow.CreateCell(column.Ordinal);
                    headerCell.SetCellValue(column.ColumnName);
                    headerCell.CellStyle = style;
                }
                rowIndex++;
            }
            foreach (DataRow row in table.Rows)
            {
                var dataRow = sheet.CreateRow(rowIndex++);

                foreach (DataColumn column in table.Columns)
                {
                    var cell = dataRow.CreateCell(column.Ordinal);
                    cell.SetCellValue(row[column].ToString());
                    cell.CellStyle = style;
                }
            }
        }

        #endregion

        #region--insert image

        private void InsertImage(ISheet sheet, Stream imageStream, int top, int bottom, int left, int right)
        {
            var bytes = StreamToBytes(imageStream);
            int pictureIdx = workbook.AddPicture(bytes, PictureType.JPEG);
            // Create the drawing patriarch.  This is the top level container for all shapes. 
            var patriarch = (HSSFPatriarch) sheet.CreateDrawingPatriarch();
            ////add a picture
            HSSFClientAnchor anchor = new HSSFClientAnchor();
            anchor.Col1 = left/cellPiexWidth;
            anchor.Row1 = top/cellPiexHeight;
            anchor.Col2 = right/cellPiexWidth;
            anchor.Row2 = bottom/cellPiexHeight;
            anchor.Dx1 = GetAnchorX(left%cellPiexWidth);
            anchor.Dy1 = GetAnchorY(top%cellPiexHeight);
            anchor.Dx2 = GetAnchorX(right%cellPiexWidth);
            anchor.Dy2 = GetAnchorY(bottom%cellPiexHeight);
            //HSSFClientAnchor anchor = new HSSFClientAnchor(500, 200, 1023, 100, 0
            //    , 0, 7, 9);
            patriarch.CreatePicture(anchor, pictureIdx);

            //pict.Resize();
        }

        public void InsertImage(Stream imageStream, int width, int height, int marginTop, int marginLeft)
        {
            InsertImage(currentSheet, imageStream, marginTop, marginTop + height, marginLeft, marginLeft + width);
        }


        public void InsertImage(Stream imageStream, int marginTop, int marginLeft)
        {                  
            Image image = Image.FromStream(imageStream);
            InsertImage(currentSheet, imageStream, marginTop, marginTop + image.Height, marginLeft,
                marginLeft + image.Height);
        }



        private int GetAnchorX(int px)
        {
            return (int)Math.Round(1023d / cellPiexWidth * px);
        }

        private int GetAnchorY(int px)
        {
            return (int)Math.Round(255d / cellPiexHeight * px);
        }       

        private byte[] StreamToBytes(Stream stream)
        {
            byte[] bytes = new byte[stream.Length];
            stream.Position = 0; //置于流开始位置
            stream.Read(bytes, 0, bytes.Length);
            stream.Seek(0, SeekOrigin.Begin);
            return bytes;
        }

        private Stream BytesToStream(byte[] bytes)
        {
            Stream stream = new MemoryStream(bytes);
            return stream;
        }

        #endregion

        #region--insert textbox

        public void InsertTextBox(string text, int top, int bottom, int left, int right)
        {
            var patriarch = (HSSFPatriarch)currentSheet.CreateDrawingPatriarch();
            HSSFClientAnchor anchor = new HSSFClientAnchor();            
            anchor.Col1 = left / cellPiexWidth;
            anchor.Row1 = top / cellPiexHeight;
            anchor.Col2 = right / cellPiexWidth;
            anchor.Row2 = bottom / cellPiexHeight;
            anchor.Dx1 = GetAnchorX(left % cellPiexWidth);
            anchor.Dy1 = GetAnchorY(top % cellPiexHeight);
            anchor.Dx2 = GetAnchorX(right % cellPiexWidth);
            anchor.Dy2 = GetAnchorY(bottom % cellPiexHeight);

            var textbox = patriarch.CreateTextbox(anchor);
            textbox.String = new HSSFRichTextString(text);            
            textbox.IsNoFill = true;
            textbox.LineStyle = LineStyle.None;
            
        }

        #endregion

        #region--find cell / find all

        private ICell Find(ISheet sheet, string text)
        {
            for (int rowIndex = 0; rowIndex < sheet.LastRowNum; rowIndex++)
            {
                IRow row = GetRow(rowIndex);
                for (int cellIndex = 0; cellIndex < row.LastCellNum; cellIndex++)
                {
                    ICell cell = GetCell(row, cellIndex);
                    if (cell.StringCellValue.Equals(text))
                    {
                        return cell;
                    }
                }
            }
            return null;
        }

        private ICell Find(CellRangeAddress range, string text)
        {
            for (int rowIndex = range.FirstRow; rowIndex <= range.LastRow; rowIndex++)
            {
                IRow row = GetRow(rowIndex);
                for (int cellIndex = range.FirstColumn; cellIndex <= range.LastColumn; cellIndex++)
                {
                    ICell cell = GetCell(row, cellIndex);
                    if (cell.StringCellValue.Equals(text))
                    {
                        return cell;
                    }
                }
            }
            return null;
        }


        private List<string> FindAllText(ISheet sheet, string pattern)
        {
            List<string> labels = new List<string>();
            Regex labelRegex = new Regex(pattern);
            for (int rIndex = sheet.FirstRowNum; rIndex < sheet.LastRowNum; rIndex++)
            {
                IRow row = GetRow(rIndex);
                for (int cIndex = row.FirstCellNum; cIndex < row.LastCellNum; cIndex++)
                {
                    ICell cell = GetCell(row, cIndex);
                    string strValue = cell.StringCellValue;
                    if (labelRegex.IsMatch(strValue))
                    {
                        MatchCollection matchCollection = labelRegex.Matches(strValue);
                        foreach (Match match in matchCollection)
                        {
                            labels.Add(match.Value);
                        }
                    }
                }
            }
            return labels;
        }

        private List<string> FindAllText(CellRangeAddress range, string pattern)
        {
            return FindAll(range.FirstRow, range.LastRow, range.FirstColumn, range.LastColumn, pattern);
        }


        private List<string> FindAll(int firstRow, int lastRow, int firstColumn, int lastColumn, string pattern)
        {
            List<string> labels = new List<string>();
            Regex labelRegex = new Regex(pattern);
            for (int rowIndex = firstRow; rowIndex <= lastRow; rowIndex++)
            {
                IRow row = GetRow(rowIndex);
                for (int cellIndex = firstColumn; cellIndex <= lastColumn; cellIndex++)
                {
                    ICell cell = GetCell(row, cellIndex);
                    string strValue = cell.StringCellValue;
                    if (labelRegex.IsMatch(strValue))
                    {
                        MatchCollection matchCollection = labelRegex.Matches(strValue);
                        foreach (Match match in matchCollection)
                        {
                            labels.Add(match.Value);
                        }
                    }
                }
            }
            return labels;
        }

        #endregion

        #region--find regin / merge region

        private CellRangeAddress FindRange(string rangeName)
        {
            var name = workbook.GetName(rangeName);
            var range = CellRangeAddress.ValueOf(name.RefersToFormula);
            return range;
        }


        public void MergedRegion(int firstRow, int lastRow, int firstCol, int lastCol)
        {
            currentSheet.AddMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
        }

        #endregion

        #region -- set border

        public void SetBorder(int firstRow, int lastRow, int firstCol, int lastCol)
        {
            for (int rowIndex = firstRow; rowIndex < lastRow; rowIndex++)
            {
                var row = HSSFCellUtil.GetRow(rowIndex, currentSheet);
                for (int cellIndex = firstCol; cellIndex < lastCol; cellIndex++)
                {
                    var cell = HSSFCellUtil.GetCell(row, cellIndex);
                    cell.CellStyle = GetThinBDRStyle();
                }
            }
        }

        private void SetBorderLeft(int firstRow, int lastRow, int firstCol, int lastCol)
        {
            HSSFRegionUtil.SetBorderLeft(BorderStyle.THIN, new CellRangeAddress(firstRow, lastRow, firstCol, lastCol),
                currentSheet, workbook);
        }

        private void SetBorderRight(int firstRow, int lastRow, int firstCol, int lastCol)
        {
            HSSFRegionUtil.SetBorderRight(BorderStyle.THIN, new CellRangeAddress(firstRow, lastRow, firstCol, lastCol),
                currentSheet, workbook);
        }

        private void SetBorderBottom(int firstRow, int lastRow, int firstCol, int lastCol)
        {
            HSSFRegionUtil.SetBorderTop(BorderStyle.THIN, new CellRangeAddress(firstRow, lastRow, firstCol, lastCol),
                currentSheet, workbook);
        }

        private void SetBorderTop(int firstRow, int lastRow, int firstCol, int lastCol)
        {
            HSSFRegionUtil.SetBorderTop(BorderStyle.THIN, new CellRangeAddress(firstRow, lastRow, firstCol, lastCol),
                currentSheet, workbook);
        }

        private ICellStyle GetThinBDRStyle()
        {
            ICellStyle style = workbook.CreateCellStyle();
            style.BorderRight = BorderStyle.THIN;
            style.BorderBottom = BorderStyle.THIN;
            style.BorderLeft = BorderStyle.THIN;
            style.BorderTop = BorderStyle.THIN;
            return style;
        }

        private void SetCellBorder(ISheet sheet, int firstRow, int lastRow, int firstCol, int lastCol)
        {

            for (int rowIndex = 0; rowIndex < sheet.LastRowNum; rowIndex++)
            {
                IRow row = sheet.GetRow(rowIndex);
                for (int cellIndex = 0; cellIndex < row.LastCellNum; cellIndex++)
                {
                    ICell cell = row.GetCell(cellIndex);
                    cell.CellStyle = GetThinBDRStyle();
                }
            }
        }

        #endregion

        #region--set style

        private ICellStyle CreateStyle()
        {
            ICellStyle style = workbook.CreateCellStyle();
            return style;
        }

        /// <summary>
        /// 设置单元居中
        /// </summary>
        /// <param name="row"></param>
        /// <param name="col"></param>
        public void SetCellCenter(int row, int col)
        {
            ICell cell = GetCell(row, col);
            ICellStyle style = cell.CellStyle;
            style.Alignment = HorizontalAlignment.CENTER;
            cell.CellStyle = style;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="fontHeight"></param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        public void SetCellFont(int fontHeight, int row, int col)
        {
            ICell cell = GetCell(row, col);
            ICellStyle style = cell.CellStyle;
            IFont font = workbook.CreateFont();
            font.FontHeight = (short) fontHeight;
            style.SetFont(font);
            cell.CellStyle = style;

        }

        public void SetFont(int firstRow, int lastRow, int firstCol, int lastCol)
        {
            for (int rowIndex = firstRow; rowIndex < lastRow; rowIndex++)
            {
                var row = HSSFCellUtil.GetRow(rowIndex, currentSheet);
                for (int cellIndex = firstCol; cellIndex < lastCol; cellIndex++)
                {
                    var cell = HSSFCellUtil.GetCell(row, cellIndex);
                    ICellStyle style = workbook.CreateCellStyle();
                    IFont font = workbook.CreateFont();
                    font.Boldweight = (short)FontBoldWeight.BOLD;
                    style.SetFont(font);
                    cell.CellStyle = style;
                }
            }
        }
        #endregion

     

        #region-- get row/cell

        /// <summary>
        /// 获取单元对象
        /// </summary>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        private ICell GetCell(int row, int col)
        {
            IRow curRow = GetRow(row);
            return CellUtil.GetCell(curRow, col);
        }

        /// <summary>
        /// 获取单元对象
        /// </summary>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        private ICell GetCell(IRow row, int col)
        {
            return CellUtil.GetCell(row, col);
        }

        /// <summary>
        /// 获取行对象
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        private IRow GetRow(int row)
        {
            return CellUtil.GetRow(row, currentSheet);
        }

        #endregion

        #region--debug

        [ConditionalAttribute("DEBUG")]
        public void PrintCurrentSheet()
        {
            for (int rowIndex = 0; rowIndex < currentSheet.LastRowNum; rowIndex++)
            {
                IRow row = currentSheet.GetRow(rowIndex);
                for (int cellIndex = 0; cellIndex < row.LastCellNum; cellIndex++)
                {
                    Console.WriteLine(row.GetCell(cellIndex).StringCellValue);
                }
            }
        }

        #endregion

        #region DataTable helper    

        /// <summary>
        /// 移动行
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="fromRowIndex"></param>
        /// <param name="endRowIndex"></param>
        /// <param name="n"></param>
        private void ShiftRows(ISheet sheet, int fromRowIndex, int endRowIndex, int n)
        {
            sheet.ShiftRows(fromRowIndex, endRowIndex, n, false, true);
        }

        /// <summary>
        /// 拷贝行
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="sourceRowIndex"></param>
        /// <param name="formRowIndex"></param>
        /// <param name="n"></param>
        private void CopyRows(ISheet sheet, int sourceRowIndex, int formRowIndex, int n)
        {
            for (int i = formRowIndex; i < formRowIndex + n; i++)
            {
                SheetUtil.CopyRow(sheet, sourceRowIndex, i);
            }
        }

        /// <summary>
        /// 插入行
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="fromRowIndex"></param>
        /// <param name="n"></param>
        //private void InsertRows(ISheet sheet, int fromRowIndex, int n,ICell style)
        //{
        //    for (int rowIndex = fromRowIndex; rowIndex < fromRowIndex + n; rowIndex++)
        //    {
        //        IRow rowSource = sheet.GetRow(rowIndex + n);
        //        IRow rowInsert = sheet.CreateRow(rowIndex);
        //        rowInsert.Height = rowSource.Height;
        //        for (int colIndex = 0; colIndex < rowSource.LastCellNum; colIndex++)
        //        {
        //            ICell cellSource = rowSource.GetCell(colIndex);
        //            ICell cellInsert = rowInsert.CreateCell(colIndex);
        //            if (cellSource != null)
        //            {
        //                cellInsert.CellStyle = cellSource.CellStyle;
        //            }
        //        }
        //    }
        //}

        #endregion



    }
}
