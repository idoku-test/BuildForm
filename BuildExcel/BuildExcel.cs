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
using System.Threading.Tasks;

namespace BuildExcel
{
    public class BuildExcel 
    {
        public BuildExcel()
        {
            this.workbook = new HSSFWorkbook();
            currentSheet = (HSSFSheet)this.workbook.CreateSheet("sheet1");
            this.workbook.CreateSheet("sheet2");
            this.workbook.CreateSheet("sheet3");

        }



        public BuildExcel(Stream fileStream)
        {
            workbook = new HSSFWorkbook(fileStream);
            this.currentSheet = (HSSFSheet)this.workbook.GetSheetAt(0);
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
            currentSheet = (HSSFSheet)workbook.GetSheet(sheetName);          
        }
        #endregion

        #region public
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
        
        public void Replace(string what, string replacement)
        {
            ICell cell = FindFirstCell(currentSheet, what);
            if (cell != null)
            {
                cell.SetCellValue(replacement);
            }
        }

        public void ReplaceInsertTable(string what, DataTable table)
        {
            ICell cell = FindFirstCell(currentSheet, what);
            if (cell != null)
            {
                int sheetIndex = workbook.GetSheetIndex(currentSheet);
                int rowCount = table.Rows.Count;
                int rowIndex = cell.RowIndex;
                InsertRows(currentSheet, rowIndex + 1, rowCount - 1);//remove what row
                InsertDataTable(table, sheetIndex, rowIndex, false);                
            }
        }

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

        public void MergedRegion(int firstRow, int lastRow, int firstCol, int lastCol)
        {
            currentSheet.AddMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
        }

        #endregion

    

     


        #region DataTable helper

        

        public void InsertDataTable(DataTable table)
        {
            InsertDataTable(table, 0, 0, true);
        }

        public void InsertDataTable(DataTable table, int sheetIndex,int rowIndex,bool hasHeader)
        {
            string sheetName = workbook.GetSheetName(sheetIndex);
            var sheet = workbook.GetSheet(sheetName);
            var style = GetThinBDRStyle();
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

        /// <summary>
        /// 插入行
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="fromRowIndex"></param>
        /// <param name="rowCount"></param>
        private void InsertRows(ISheet sheet, int fromRowIndex, int rowCount)
        {
            sheet.ShiftRows(fromRowIndex, sheet.LastRowNum, rowCount, false, true);
            for (int rowIndex = fromRowIndex; rowIndex < fromRowIndex+rowCount; rowIndex++)
            {
                IRow rowSource = sheet.GetRow(rowIndex + rowCount);
                IRow rowInsert = sheet.CreateRow(rowIndex);
                rowInsert.Height = rowSource.Height;
                for (int colIndex = 0; colIndex < rowSource.LastCellNum; colIndex++)
                {
                    ICell cellSource = rowSource.GetCell(colIndex);
                    ICell cellInsert = rowInsert.CreateCell(colIndex);
                    if (cellSource != null)
                    {
                        cellInsert.CellStyle = cellSource.CellStyle;
                    }
                }
            }
        }
        #endregion

     


    

        #region 
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

        private ICell FindFirstCell(ISheet sheet, string text)
        {                      
            for (int rowIndex = 0; rowIndex < sheet.LastRowNum; rowIndex++)
            {
                IRow row = sheet.GetRow(rowIndex);
                for (int cellIndex = 0; cellIndex < row.LastCellNum; cellIndex++)
                {
                    ICell cell = row.GetCell(cellIndex);
                    if (cell.StringCellValue.Equals(text))
                    {
                        return cell;
                    }
                }
            }
            return null;
        }
        #endregion

    }
}
