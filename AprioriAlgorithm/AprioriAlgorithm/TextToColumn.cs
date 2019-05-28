using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace AprioriAlgorithm
{
    class TextToColumn
    {
        //Veri setinin sütunlara bölünmesini sağlayan metod. Parçalama işlemini virgülü baz alarak yapmaktadır.
        public Worksheet textToColumn(string[] dataSetString, Worksheet currentSheet, Application currentApp, Workbook currentBook)
        {
            //NEW WORKSHEET
            // Kullanıcı veri setinin tekrar kullanılabilirliği açısından verileri  yedekleyip yeni bir Excel çalışma sayfası açılarak parçalama işlemi yapılıyor.
            if (currentApp.Application.Sheets.Count < 2)
            {
                currentSheet = (Excel.Worksheet)currentBook.Worksheets.Add();
            }
            else
            {
                currentSheet = currentApp.Worksheets[2];
            }
            currentSheet.Name = "Data Set";
            int elementCount = 0;
            for (int i = 0; i < dataSetString.Length; i++)
            {
                char spl = ',';
                string[] parca;
                parca = dataSetString[i].Split(spl);
                elementCount = 0;
                for (char j = 'A'; j < 'A' + parca.Length; j++)
                {
                    Range print = currentSheet.Range["" + j + (i + 1)];
                    print.Value = parca[elementCount];
                    print.EntireColumn.AutoFit();
                    elementCount++;
                }
            }
            return currentSheet;
        }
    }
}
