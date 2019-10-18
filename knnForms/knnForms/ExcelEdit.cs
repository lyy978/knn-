using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace knnForms
{
    public class ExcelEdit
    {
        public string mFilename;
        public Microsoft.Office.Interop.Excel.Application app;
        public Microsoft.Office.Interop.Excel.Workbooks wbs;
        public Microsoft.Office.Interop.Excel.Workbook wb;
        public Microsoft.Office.Interop.Excel.Worksheets wss;
        public Microsoft.Office.Interop.Excel.Worksheet ws;

        public void Create()//create a Microsoft.Office.Interop.Excel object
        {
            app = new Microsoft.Office.Interop.Excel.Application();
            wbs = app.Workbooks;
            wb = wbs.Add(true);
        }
        public void Open(string FileName)//open a Microsoft.Office.Interop.Excel file
        {
            app = new Microsoft.Office.Interop.Excel.Application();
            wbs = app.Workbooks;
            wb = wbs.Add(FileName);
            mFilename = FileName;
        }
        public Microsoft.Office.Interop.Excel.Worksheet GetSheet(string SheetName)   //open a sheet
        {
            Microsoft.Office.Interop.Excel.Worksheet s = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets[SheetName];
            return s;
        }
        public Microsoft.Office.Interop.Excel.Worksheet AddSheet(string SheetName)  //add a worksheet
        {
            Microsoft.Office.Interop.Excel.Worksheet s = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            s.Name = SheetName;
            return s;
        }
        public void Close()
        //destory Microsoft.Office.Interop.Excel object
        {
            wb.Save();
            wb.Close(Type.Missing, Type.Missing, Type.Missing);
            wbs.Close();
            app.Quit();
            wb = null;
            wbs = null;
            app = null;
            GC.Collect();
        }
        public bool Save()  //save the file
        {
            if (mFilename == "")
            {
                return false;
            }
            else
            {
                try
                {
                    wb.Save();
                    return true;
                }

                catch (Exception ex)
                {
                    return false;
                }
            }
        }
        public void SetCellValue(Microsoft.Office.Interop.Excel.Worksheet ws, int x, int y, object value)
        //ws：setting value worksheet     X row Y column     
        {
            ws.Cells[x, y] = value;

        }
        public void SetCellValue(string ws, int x, int y, object value)
        {

            GetSheet(ws).Cells[x, y] = value;
        }

        public List<string> ColumnDB = new List<string>();
        public List<int> trainX1 = new List<int>();
        public List<int> trainX2 = new List<int>();
        public List<int> trainX3 = new List<int>();
        public List<double> trainX4 = new List<double>();
        public List<double> trainX5 = new List<double>();
        public List<double> trainY1 = new List<double>();
        public List<double> trainY2 = new List<double>();
        public List<double> trainY3 = new List<double>();


        public void getColumnInt(ExcelEdit ed)
        {

            Excel.Worksheet worksheet = (Excel.Worksheet)ed.GetSheet("sheet1");  //get worksheet

            int rows = worksheet.UsedRange.Rows.Count;                           //get count of rows  of worksheet
            int columns = worksheet.UsedRange.Columns.Count;                     //get count of columns of worksheet 
            Console.WriteLine("which column");
            int column = Convert.ToInt16(Console.ReadLine());
            int m = 0;
            // read datas by column
            for (int i = 2; i <= rows; i++)
            {
                int temp;
                string a = (worksheet.Cells[i, column]).Text.ToString();
                temp = Convert.ToInt32(a);
                ColumnDB.Add(a);
                trainX1.Add(temp);
                m++;
            }

            //  Console.WriteLine("{0}", rows);

            Console.ReadLine();
        }

        public void getTestnum(ExcelEdit ed)                                                 //获取测试集数据
        {
            Excel.Worksheet worksheet = (Excel.Worksheet)ed.GetSheet("sheet1");  //get worksheet

            int rows = worksheet.UsedRange.Rows.Count;                         //get count of rows  of worksheet 
            int columns = worksheet.UsedRange.Columns.Count;                    //get count of columns of worksheet 




            //read by rows
            for (int i = 2; i <= rows; i++)    //read column 1 data 
            {
                int temp;
                string a = (worksheet.Cells[i, 1]).Text.ToString();
                temp = Convert.ToInt32(a);
                ColumnDB.Add(a);
                trainX1.Add(temp);
            }

            for (int i = 2; i <= rows; i++)    //column 2
            {
                int temp;
                string a = (worksheet.Cells[i, 2]).Text.ToString();
                temp = Convert.ToInt32(a);
                ColumnDB.Add(a);
                trainX2.Add(temp);
            }

            for (int i = 2; i <= rows; i++)    //column 3
            {
                int temp;
                string a = (worksheet.Cells[i, 3]).Text.ToString();
                temp = Convert.ToInt32(a);
                ColumnDB.Add(a);
                trainX3.Add(temp);
            }

            for (int i = 2; i <= rows; i++)    //column 4
            {
                double temp;
                string a = (worksheet.Cells[i, 4]).Text.ToString();
                temp = Convert.ToDouble(a);
                ColumnDB.Add(a);
                trainX4.Add(temp);
            }

            for (int i = 2; i <= rows; i++)    //column 5
            {
                double temp;
                string a = (worksheet.Cells[i, 5]).Text.ToString();
                temp = Convert.ToDouble(a);
                ColumnDB.Add(a);
                trainX5.Add(temp);
            }

            for (int i = 2; i <= rows; i++)    //column 6
            {
                double temp;
                string a = (worksheet.Cells[i, 6]).Text.ToString();
                trainY1.Add(Convert.ToDouble(a));
            }

            for (int i = 2; i <= rows; i++)    //column 7
            {
                double temp;
                string a = (worksheet.Cells[i, 7]).Text.ToString();
                trainY2.Add(Convert.ToDouble(a));
            }

            for (int i = 2; i <= rows; i++)    //column 8
            {
                double temp;
                string a = (worksheet.Cells[i, 8]).Text.ToString();
                trainY3.Add(Convert.ToDouble(a));
            }
        }
    }
}
