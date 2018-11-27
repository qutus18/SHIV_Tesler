using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ShivExcelLogging
{
    public partial class UpdateDataExcel : Form
    {
        Excel.Application myExcel = new Excel.Application();
        Stream stream;
        SpreadsheetDocument spreadsheetDocument;
        SharedStringTablePart shareStringpart;
        public UpdateDataExcel()
        {
            InitializeComponent();
        }

        private void btnOpen_Click(object sender, EventArgs e)
        {
            string strDoc = @"‪C:\Users\Admin\Desktop\DS.xlsx";
            stream = File.Open(strDoc, FileMode.Open);
            spreadsheetDocument = SpreadsheetDocument.Open(stream, true);
            if (spreadsheetDocument.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
            {
                shareStringpart = spreadsheetDocument.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
            }
            else
            {
                shareStringpart = spreadsheetDocument.WorkbookPart.AddNewPart<SharedStringTablePart>();
            }
        }

        private void UpdateDataExcel_FormClosing(object sender, FormClosingEventArgs e)
        {
            stream.Close();
        }
    }
}
