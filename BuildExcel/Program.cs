using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using NPOI.HSSF.UserModel;
using System.IO;
using System.Data;
using System.Drawing;
using System.Net.Mime;
using NPOI.SS.UserModel;

namespace BuildExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            //BuildExcel excel = new BuildExcel();
            //FileStream file = new FileStream(@"Excel/template3.xls", FileMode.Open, FileAccess.ReadWrite);
            //BuildExcel excel = new BuildExcel(file);
            //excel.InsertTextBox("附件一：估价对象侨苑（二期）地理位置图", 0, 1, 2, 4);
            //excel.SaveAs("d.xls");
            //Stream ms = excel.GetStream();
            //FileStream saveTo = new FileStream("d.xls", FileMode.Create);
            //ms.CopyTo(saveTo);

            //saveTo.Close();
            //file.Close();

            FontTest();

            Console.Read();
        }

        private static void FontTest()
        {
            BuildExcel excel = new BuildExcel();
            DataTable table = CreateTable(3, 8);
            excel.InsertTable(table, 0, 0);
            excel.SetFont(0,1,0,5);
            Stream ms = excel.GetStream();
            FileStream file = new FileStream("d.xls", FileMode.Create);
            ms.CopyTo(file);
            excel.GetStream();
        }

        private static void ImageTest()
        {
            FileStream file = new FileStream(@"Excel/template3.xls", FileMode.Open, FileAccess.ReadWrite);
            BuildExcel excel = new BuildExcel(file);

            var image = Image.FromFile(@"Image/201508211237340.jpg");
            var image2 = Image.FromFile(@"Image/201508211237341.jpg");
            //excel.ReplaceInRange("a", "b", "temp");
            MemoryStream ims = new MemoryStream();
            image.Save(ims, System.Drawing.Imaging.ImageFormat.Jpeg);
            ims.Position = 0;

            MemoryStream ims2 = new MemoryStream();
            image2.Save(ims2, System.Drawing.Imaging.ImageFormat.Jpeg);
            ims2.Position = 0;
            //excel.InsertImage();
            var spacingWidth = CmToPx(1.5);
            var spacingHeight = CmToPx(2.1);

            var leftMargin = CmToPx(1);
            var topMargin = CmToPx(1.8);

            var photoHeight = CmToPx(8.5);
            var photoWidth = CmToPx(6.4);

            excel.InsertImage(ims, photoWidth, photoHeight, CmToPx(1.8), CmToPx(1));
            excel.InsertImage(ims2, photoWidth, photoHeight, CmToPx(1.8), CmToPx(1 + 2) + photoWidth);
            Stream ms = excel.GetStream();
            FileStream saveTo = new FileStream("d.xls", FileMode.Create);
            ms.CopyTo(saveTo);

            saveTo.Close();
            file.Close();
        }

        private static int CmToPx(double cm, int dpi = 96)
        {
            return (int)Math.Floor(cm / 2.54 * dpi);
        }
        private static int CmToPt(double cm, int dpi = 96)
        {
            return (int)Math.Floor(cm * 72 / dpi * 2.54);
        }

        private static void BookMarkTest()
        {
            FileStream file = new FileStream(@"Excel/template2.xls", FileMode.Open, FileAccess.ReadWrite);
            BuildExcel excel = new BuildExcel(file);
            var list = excel.GetBookmarks();
            foreach (var item in list)
            {
                Console.Write(item + " ");
            }
        }

        private static void StatsTest()
        {
            BuildExcel excel = new BuildExcel();
            excel.InsertText("业务统计", 0, 0);
            excel.SetCellCenter(0, 0);
            excel.SetCellFont(15*15, 0, 0);
            excel.MergedRegion(0, 0, 0, 4);
            excel.InsertText("统计时间", 1, 3);
            excel.InsertText("2015/5/1-2015/6/1", 1, 4);
            DataTable table = CreateTable(5, 5);
            excel.InsertTable(table, 2, 0);
            Stream ms = excel.GetStream();
            FileStream file = new FileStream("c.xls", FileMode.Create);
            ms.CopyTo(file);
            file.Close();
        }


        public static void DictionaryTable()
        {
            Dictionary<string, string> dic = new Dictionary<string, string>() { 
            { "楼盘名称","123" },
            { "楼盘地址","321" },
            {"楼栋名称","123"},
            {"楼栋地址","123"},
            {"楼栋号","123"}
            };

            var dt = DictionaryToTable(dic, 3);
            BuildExcel excel = new BuildExcel();
            //excel.InsertTable(dt, 0, true);
            Stream ms = excel.GetStream();
            FileStream file = new FileStream("a.xls", FileMode.Create);
            ms.CopyTo(file);
            file.Close();
        }

        public static DataTable DictionaryToTable(Dictionary<string, string> dictionary, int dicToColumnNum)
        {
            int colNum = dicToColumnNum * 2;
            int rowNum = dictionary.Count() / dicToColumnNum;
            DataTable dt = CreateTable(rowNum, colNum);

            if (dicToColumnNum <= 0)
            {
                throw new ArgumentOutOfRangeException("dicToColumnNum", "字典展开不能小于等于0.");
            }
            else
            {
                for (int i = 0, col = 0, row = 0; i < dictionary.Count(); i++)
                {
                    var item = dictionary.ElementAt(i);
                    col = 2 * (i % dicToColumnNum);
                    row = i / dicToColumnNum;
                    dt.Rows[row][col] = item.Key;
                    dt.Rows[row][col + 1] = item.Value;
                }
            }
            return dt;
        }

        private static DataTable CreateTable(int rowNums, int colNums)
        {
            DataTable table = new DataTable();
            for (int i = 0; i < colNums; i++)
            {
                DataColumn dc = new DataColumn();
                table.Columns.Add(dc);
            }
            for (int j = 0; j <= rowNums; j++)
            {
                DataRow row = table.NewRow();
                for (int r = 0; r < colNums; r++)
                {
                    row[r] = r + " " + j;
                }
                table.Rows.Add(row);
            }
            
            return table;
        }

        public static void InsertDataTable() {

            Dictionary<string, string> dic = new Dictionary<string, string>() { 
            { "楼盘名称","123" },
            { "楼盘地址","321" },
            {"楼栋名称","123"},
            {"楼栋地址","123"},
            {"楼栋号","123"}
            };

            var dt = DictionaryToTable(dic, 4);
            FileStream file = new FileStream(@"Excel/template.xls", FileMode.Open, FileAccess.ReadWrite);
            BuildExcel excel = new BuildExcel(file);
         
            excel.ReplaceInsertTable("《查勘信息》", dt);
            
            excel.SetBorder(0, 0, 1, 1);
            Stream ms = excel.GetStream();                      
            FileStream saveTo = new FileStream("b.xls", FileMode.Create);
            ms.CopyTo(saveTo);
            saveTo.Close();
            file.Close();
        }

        public static void TransDataTable()
        {
            DataTable dt = new DataTable("cart");
            DataColumn dc1 = new DataColumn("prizename", Type.GetType("System.String"));
            DataColumn dc2 = new DataColumn("point", Type.GetType("System.Int16"));
            DataColumn dc3 = new DataColumn("number", Type.GetType("System.Int16"));
            DataColumn dc4 = new DataColumn("totalpoint", Type.GetType("System.Int64"));
            DataColumn dc5 = new DataColumn("prizeid", Type.GetType("System.String"));
            DataColumn dc6 = new DataColumn("prizeid2", Type.GetType("System.String"));
            DataColumn dc7 = new DataColumn("prizeid3", Type.GetType("System.String"));
            DataColumn dc8 = new DataColumn("prizeid4", Type.GetType("System.String"));
            dt.Columns.Add(dc1);
            dt.Columns.Add(dc2);
            dt.Columns.Add(dc3);
            dt.Columns.Add(dc4);
            dt.Columns.Add(dc5);
            dt.Columns.Add(dc6);
            dt.Columns.Add(dc7);
            dt.Columns.Add(dc8);
            //以上代码完成了DataTable的构架，但是里面是没有任何数据的
            for (int i = 0; i < 10; i++)
            {
                DataRow dr = dt.NewRow();
                dr["prizename"] = "娃娃";
                dr["point"] = 10;
                dr["number"] = 1;
                dr["totalpoint"] = 10;
                dr["prizeid"] = "001";
                dr["prizeid2"] = "002";
                dr["prizeid3"] = "003";
                dr["prizeid4"] = "004";
                dt.Rows.Add(dr);
            }

            BuildExcel excel = new BuildExcel();
          
            //excel.InsertTable(dt,0,true);
            Stream ms = excel.GetStream();
            FileStream file = new FileStream("b.xls", FileMode.Create);
            ms.CopyTo(file);
            file.Close();

        }
    }
}
