using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog sfd = new SaveFileDialog() {Filter = "Excel Workbook|*.xlsx", ValidateNames = true})
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
                    Workbook wb = app.Workbooks.Add(XlSheetType.xlWorksheet);
                    
                    Worksheet ws = (Worksheet) app.ActiveSheet;
                    ws.Name = "sqsq";
                    Worksheet newWorksheet;
                    newWorksheet = (Worksheet) wb.Worksheets.Add(Type.Missing, ws, Type.Missing, Type.Missing);
                    newWorksheet.Name = "dddddd";
                    ws.Activate();
                    app.Visible = false;
                    ws.Cells[1, 1] = "aaa";
                    ws.Cells[1, 2] = "bbb";
                    ws.Cells[1, 3] = "ccc";

                    wb.SaveAs(sfd.FileName, XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, false,false,
                        XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);
                    app.Quit();
                }
            }
        }
    }
}
