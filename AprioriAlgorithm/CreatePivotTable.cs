using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace AprioriAlgorithm
{
    class CreatePivotTable
    {
	 // Parametre olarak aktif çalışma sayfası, aktif uygulama, aktf çalışma kitabını alan ve sonucunda bir pivot tablo çizen metod. 
        public Worksheet createPivotTable(Worksheet currentSheet, Application currentApp, Workbook currentBook, int rowCount)
        {
	    // Pivot tablo yeni bir excel çalışma sayfası açılarak yazdırılıyor.
            Excel.Range oRange = currentSheet.Range["C1:D" + rowCount];
            if (currentApp.Application.Sheets.Count < 3)
            {
                currentSheet = (Excel.Worksheet)currentBook.Worksheets.Add();
            }
            else
            {
                currentSheet = currentApp.Worksheets[3];
            }
            currentSheet.Name = "Apriori Pivot Table";

            Excel.Range oRange2 = currentSheet.Cells[1, 1];
            Excel.PivotCache oPivotCache = (Excel.PivotCache)currentBook.PivotCaches().Add(Excel.XlPivotTableSourceType.xlDatabase, oRange);
            Excel.PivotTable oPivotTable = (Excel.PivotTable)currentSheet.PivotTables().Add(PivotCache: oPivotCache, TableDestination: oRange2, TableName: "Apriori Pivot Table");

	    // Çizilecek pivot tablo referans olarak işlem numarası ve ürünleri alıyor.
            // işlem numarasına göre hangi ürünler o işlem içerisinde varsa tabloda işaretleme yapıyor.
            Excel.PivotField oPivotField1 = (Excel.PivotField)oPivotTable.PivotFields("Transaction");
            oPivotField1.Orientation = Excel.XlPivotFieldOrientation.xlRowField;

            Excel.PivotField oPivotField2 = (Excel.PivotField)oPivotTable.PivotFields("Item");
            oPivotField2.Orientation = Excel.XlPivotFieldOrientation.xlColumnField;

            Excel.PivotField oPivotField = (Excel.PivotField)oPivotTable.PivotFields("Item");
            oPivotField.Orientation = Excel.XlPivotFieldOrientation.xlDataField;
            oPivotField.Function = Excel.XlConsolidationFunction.xlCount;

            return currentSheet;
        }

	// Tablo çizen metod
        public void createTable(Worksheet currentSheet)
        {
            Excel.Range last = currentSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            Excel.Range rdataSet = currentSheet.get_Range("A1", last);
            Excel.Borders border = rdataSet.Borders;
            border.LineStyle = Excel.XlLineStyle.xlContinuous;
            border.Weight = 3d;
        }
    }
}
