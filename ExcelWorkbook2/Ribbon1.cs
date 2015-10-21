using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using oledb = System.Data.OleDb;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;



namespace ExcelWorkbook2
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application _App = Globals.ThisWorkbook.Application as Excel.Application;
            Excel.Worksheet _sheet = _App.ActiveSheet as Excel.Worksheet;
             //Range("A5", "A5") as Excel.Range;
            //_r.Value2 = "Alicja w krainie czarow A5";
            DataSet1TableAdapters.v_listaKalkulacjiTableAdapter vTa = new DataSet1TableAdapters.v_listaKalkulacjiTableAdapter();
            DataSet1.v_listaKalkulacjiDataTable dt = new DataSet1.v_listaKalkulacjiDataTable();
            vTa.Fill(dt);
            // Range _r = _sheet.Range["A2"];

            //            _r.CopyFromRecordset()

            //   Globals.Arkusz1.list1.DataSource = dt;
            Globals.Arkusz1.list1.AutoSetDataBoundColumnHeaders = true;
            Globals.Arkusz1.list1.SetDataBinding(dt);//, "id", "Numer Zapyania", "Indeks XL", "Numer Rysunku", "Klient", "Wielkosc Zamowienia", "Wielkosc Produkcji", "Grubosc Materialu", "Rodzaj Materialu", "Wymiar X", "Wymiar Y", "Rewizja");
           // Globals.Arkusz1.list1.columns
            Globals.Arkusz1.list1.Disconnect();
           
            // Bind the list object to the Customers table.
            //list1.AutoSetDataBoundColumnHeaders = true;
            //list1.DataSource = dt;
            //list1.DataMember = "Kalkulacje";

        }
     
    }
}
