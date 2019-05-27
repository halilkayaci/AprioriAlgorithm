using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace AprioriAlgorithm
{
    class ResultPrint
    {
        // Parametre olarak verilen çalışma sayfasına ve sütuna, yine parametre olarak verilmiş ilişki isimlerini ve ilişkinin değerini yazdıran metod.
        public Worksheet printResult(Worksheet rcurrentSheet, double[] array, string[] arrayS, string column)
        {
            for (int i = 0; i < array.Length; i++)
            {
                Range val = rcurrentSheet.Range["" + column + (i + 2)];
                val.EntireColumn.AutoFit();
                val.EntireRow.Font.Bold = true;
                val.Value = arrayS[i] + array[i];
            }
            return rcurrentSheet;
        }
    }
}
