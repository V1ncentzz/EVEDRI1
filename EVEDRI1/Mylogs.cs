using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static Guna.UI2.Native.WinApi;

namespace EVEDRI1
{
    class Mylogs
    {
        //Workbook workbook = new Workbook();
        public void insertLog(string user, string message)
        {
            Workbook book = new Workbook();
            book.LoadFromFile(@"C:\Users\HF\Desktop\Main\Book1.xlsx");
            Worksheet sheet = book.Worksheets[2];

            int row = sheet.Rows.Length + 1;
            sheet.Range[row, 1].Value = user;
            sheet.Range[row, 2].Value = message;
            sheet.Range[row, 3].Value = DateTime.Now.ToString("MM/dd/yyyy");
            sheet.Range[row, 4].Value = DateTime.Now.ToString("hh : mm : ss : tt");
            book.SaveToFile(@"C:\Users\HF\Desktop\Main\Book1.xlsx", ExcelVersion.Version2016);
        }

        //public void insertlog(string user, string message)
        //{
        //    workbook.LoadFromFile(@"C:\Users\HF\Desktop\Main\Book1.xlsx"); //File location
        //    Worksheet sheet = workbook.Worksheets[1];
        //    int row = sheet.Rows.Length +1;
        //    sheet.Range[row, 1].Value = user;
        //    sheet.Range[row, 2].Value = message;
        //    sheet.Range[row, 3].Value = DateTime.Now.ToString("MM/dd/yyyy");
        //    sheet.Range[row, 4].Value = DateTime.Now.ToString("HH:mm:ss tt");
        //    //MessageBox.Show(DateTime.Now.ToString("hh:mm:ss tt"));
        //    workbook.SaveToFile(@"C:\Users\HF\Desktop\Main\Book1.xlsx", ExcelVersion.Version2016);
        //}
        //public void showlogs( DataGridView v)
        //{
        //    workbook.LoadFromFile(@"C:\Users\HF\Desktop\Main\Book1.xlsx"); //File location
        //    Worksheet worksheet = workbook.Worksheets[1];
        //    DataTable dt = worksheet.ExportDataTable();
        //    v.DataSource = dt;
        //}   
    }
}
