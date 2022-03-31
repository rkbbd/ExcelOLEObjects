using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace WinFormsApp2
{
    public partial class Form1 : Form
    {
        Excel.Application xlApp;
        Excel.Workbook xlWorkBook;
        Excel.Worksheet xlWorkSheet;

        object misValue = System.Reflection.Missing.Value;

        int i = 1;
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            newExcel();
        }

        private void newExcel()
        {
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            //add data
            xlWorkSheet.Cells[1, 1] = "";
            xlWorkSheet.Cells[1, 2] = "Student1";
            xlWorkSheet.Cells[1, 3] = "Student2";
            xlWorkSheet.Cells[1, 4] = "Student3";

            xlWorkSheet.Cells[2, 1] = "Term1";
            xlWorkSheet.Cells[2, 2] = "80";
            xlWorkSheet.Cells[2, 3] = "65";
            xlWorkSheet.Cells[2, 4] = "45";

            xlWorkSheet.Cells[3, 1] = "Term2";
            xlWorkSheet.Cells[3, 2] = "78";
            xlWorkSheet.Cells[3, 3] = "72";
            xlWorkSheet.Cells[3, 4] = "60";

            xlApp.Visible = true;


            xlWorkSheet.Shapes.AddOLEObject("Forms.CommandButton.1", Type.Missing, false, false, Type.Missing, Type.Missing, Type.Missing, 10, 10, 100, 30);
            Microsoft.Office.Interop.Excel.OLEObject oleShape = xlWorkSheet.OLEObjects(1);
            Microsoft.Vbe.Interop.Forms.CommandButton button = oleShape.Object;
            button.Caption = "Custom Buttom" + i;

            button.Click += Button_Click;
            i++;
        }
        private void Button_Click()
        {
            MessageBox.Show("It works!");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            openExisting();
        }

        private void openExisting()
        {
            xlApp = new Excel.Application();
            string workbookPath = @"C:\Users\USER\Documents\Book.xlsx";
            xlWorkBook = xlApp.Workbooks.Open(workbookPath,
                    0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "",
                    true, false, 0, true, false, false);

            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            xlApp.Visible = true;

            Microsoft.Office.Interop.Excel.OLEObject oleShape = xlWorkSheet.OLEObjects(1);
            Microsoft.Vbe.Interop.Forms.CommandButton button = oleShape.Object;
            button.Caption = "Custom Buttom" + i;

            button.Click += Button_Click;
            i++;
        }
    }
}
