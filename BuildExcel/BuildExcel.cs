using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using NPOI.HSSF.Record.CF;

namespace BuildExcel
{
    public class BuildExcel
    {
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


        #region 页操作

        /// <summary>
        /// 选择操作页
        /// </summary>
        /// <param name="sheetName"></param>
        public void SelectSheet(string sheetName)
        {
            currentSheet = (HSSFSheet) workbook.GetSheet(sheetName);
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

        #region --insert text

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

        //public List<string> GetBookmarks(string range)
        //{
        //    var r = FindRange(range);
        //    return FindAllTextInRange(r, @"《([^》]+)》");
        //}

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

        //private List<string> FindAllTextInRange(CellRangeAddress range, string pattern)
        //{
        //    return FindAll(range.FirstRow, range.LastRow, range.FirstColumn, range.LastColumn, pattern);
        //}

        //private List<string> FindAll(int firstRow, int lastRow, string pattern)
        //{
        //    List<string> labels = new List<string>();
        //    Regex labelRegex = new Regex(pattern);
        //    for (int rowIndex = firstRow; rowIndex <= lastRow; rowIndex++)
        //    {
        //        IRow row = GetRow(rowIndex);
        //        for (int cellIndex = row.FirstCellNum; cellIndex <= row.LastCellNum; cellIndex++)
        //        {
        //            var cell = row.GetCell(cellIndex);
        //            var strValue = cell.StringCellValue;
        //            labels.AddRange(Matches(strValue, labelRegex));
        //        }
        //    }
        //    return labels;
        //}


        //private List<string> Matches(string value, Regex regex)
        //{
        //    var labels = new List<string>();
        //    if (!regex.IsMatch(value)) return labels;
        //    var matchCollection = regex.Matches(value);
        //    labels.AddRange(from Match match in matchCollection select match.Value);
        //    return labels;
        //}


        //private List<string> FindAll(int firstRow, int lastRow, int firstColumn, int lastColumn, string pattern)
        //{
        //    List<string> labels = new List<string>();
        //    Regex labelRegex = new Regex(pattern);
        //    for (int rowIndex = firstRow; rowIndex <= lastRow; rowIndex++)
        //    {
        //        IRow row = GetRow(rowIndex);
        //        for (int cellIndex = firstColumn; cellIndex <= lastColumn; cellIndex++)
        //        {
        //            ICell cell = row.GetCell(cellIndex);
        //            string strValue = cell.StringCellValue;
        //            if (labelRegex.IsMatch(strValue))
        //            {
        //                MatchCollection matchCollection = labelRegex.Matches(strValue);
        //                foreach (Match match in matchCollection)
        //                {
        //                    labels.Add(match.Value);
        //                }
        //            }
        //        }
        //    }
        //    return labels;
        //}

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
            HSSFRegionUtil.SetBorderLeft(BorderStyle.THIN, new CellRangeAddress(firstRow, lastRow, firstCol, lastCol), currentSheet, workbook);
        }

        private void SetBorderRight(int firstRow, int lastRow, int firstCol, int lastCol)
        {
            HSSFRegionUtil.SetBorderRight(BorderStyle.THIN, new CellRangeAddress(firstRow, lastRow, firstCol, lastCol), currentSheet, workbook);
        }

        private void SetBorderBottom(int firstRow, int lastRow, int firstCol, int lastCol)
        {
            HSSFRegionUtil.SetBorderTop(BorderStyle.THIN, new CellRangeAddress(firstRow, lastRow, firstCol, lastCol), currentSheet, workbook);
        }

        private void SetBorderTop(int firstRow, int lastRow, int firstCol, int lastCol)
        {
            HSSFRegionUtil.SetBorderTop(BorderStyle.THIN, new CellRangeAddress(firstRow, lastRow, firstCol, lastCol), currentSheet, workbook);
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
        public void SetCellFont(int fontHeight,int row, int col) {
            ICell cell = GetCell(row, col);
            ICellStyle style = cell.CellStyle;
            IFont font = workbook.CreateFont();
            font.FontHeight = (short)fontHeight;
            style.SetFont(font);
            cell.CellStyle = style;

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
        private void ShiftRows(ISheet sheet, int fromRowIndex,int endRowIndex, int n)
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
        private void CopyRows(ISheet sheet, int sourceRowIndex,int formRowIndex, int n)
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
