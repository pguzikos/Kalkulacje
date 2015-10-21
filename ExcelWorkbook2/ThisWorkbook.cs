using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Tools.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace ExcelWorkbook2
{
    public partial class ThisWorkbook
    {
        private void ThisWorkbook_Startup(object sender, System.EventArgs e)
        {
            wyswietlWkomorce("A1","abra kadabra");
            Form1 frm = new Form1();
        //    frm.ShowDialog();
        }

        public void wyswietlWkomorce(String Adres, String tekst)
        {
            Excel.Worksheet activeWorksheet = ((Excel.Worksheet)Application.ActiveSheet);
            Excel.Range firstRow = activeWorksheet.get_Range(Adres);
            firstRow.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown);
            Excel.Range newFirstRow = activeWorksheet.get_Range(Adres);
            newFirstRow.Value2 = tekst;
        }

        private void ThisWorkbook_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisWorkbook_Startup);
            this.Shutdown += new System.EventHandler(ThisWorkbook_Shutdown);
        }

        #endregion

    }
}
