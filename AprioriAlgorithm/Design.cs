using System;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace AprioriAlgorithm
{
    public partial class Design
    {
        private void Design_Load(object sender, RibbonUIEventArgs e)
        {
     
        }
        
	// Uygulamanın ilk kez çalışığ çalışmadığını kontrol eden değişken.
        bool firstRun = true;
        private void btn_Uygula_Click(object sender, RibbonControlEventArgs e)
        {
	    // Uygula butonuna basıldığında aktif excel sayfası, excel kitabı uygulamaya bağlanıyor.
            Workbook currentBook = Globals.ThisAddIn.GetActiveWorkBook();
            Excel.Application currentApp = Globals.ThisAddIn.GetActiveApp();

	    // Uygulama ikinci veya daha fazla kez çalışıyorsa aktif sayfalar silinerek yeni hesaplama yapıkması sağlanıyor
            if (!firstRun)
            {
                for (int i = 0; i < 4; i++)
                {
                    ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[1]).Delete();
                }

            }
		
	    // Eklenti sekmesindeki editbox konrol edilerek işlemlere başlanıyor.
            if (SupportValue.Text != "" && SupportValue.Text !="0,00")
            {
                try
                {
                    firstRun = false;
                    Worksheet currentSheetOri = Globals.ThisAddIn.GetActiveWorkSheet();
                    double supportValue = Convert.ToDouble(SupportValue.Text);
                    int rowCount = currentSheetOri.UsedRange.Rows.Count;
                    currentSheetOri.Name = "Orijinal Data Set";
                   
		   // Sayfa yüklendiğinde "A1" hücresine "Date,Time,Transaction,Item" yazılıyor.
                    Excel.Range load = currentSheetOri.Range["A1"];
                    load.Value = "Date,Time,Transaction,Item";
                    Range dataSet = currentSheetOri.Range["A1:A" + (rowCount)];
                    Object[,] originalData = (dataSet.Value2);
                    int elementCount = 0;
		    // Veri setinin tamamı excel sayfasından alınıp bir diziye aktarılıyor.
                    string[] dataSetString = new string[rowCount];
                    foreach (object item in originalData)
                    {
                        dataSetString[elementCount] = Convert.ToString(item.ToString());
                        elementCount++;
                    }

                    Worksheet currentSheet = Globals.ThisAddIn.GetActiveWorkSheet();
	            // Tek satır olan veri "," e göre parçalanarak 4 sütuna cevriliyor. 
                    TextToColumn t2c = new TextToColumn();
                    currentSheet = t2c.textToColumn(dataSetString, currentSheetOri, currentApp, currentBook);

                    //CREATE PIVOT TABLE
		   // Veri setindeki değerler ile bir pivot tablo oluşturuluyor. 
                    CreatePivotTable cPT = new CreatePivotTable();
                    currentSheet = cPT.createPivotTable(currentSheet, currentApp, currentBook, rowCount);
                    //NEW WORKSHEET
		   // Yeni bir ecxel çalışma sayfası açılıyor
                    Worksheet ecurrentSheet = Globals.ThisAddIn.GetActiveWorkSheet();
                    if (currentApp.Application.Sheets.Count < 4)
                    {
                        ecurrentSheet = (Excel.Worksheet)currentBook.Worksheets.Add();
                    }
                    else
                    {
                        ecurrentSheet = currentApp.Worksheets[4];
                    }
                    ecurrentSheet.Name = "Apriori Algorithm";
                    //CALC SUPPORT VALUE
		   // Bütün ürünlerin tek tek destek değerleri hesaplanıyor. Yeni açılan excel sayfasında listeleniyor
                    int erowCount = ecurrentSheet.UsedRange.Rows.Count;
                    int rowCount2 = currentSheet.UsedRange.Rows.Count;
                    var columnCount = currentSheet.UsedRange.Columns.Count;
                    Range cTotal = currentSheet.Range["A" + (rowCount2 - 1)];
                    int total = Convert.ToInt32(cTotal.Value);
                    Range items = ecurrentSheet.Range["A" + (erowCount)];
                    items.EntireRow.Font.Bold = true;
                    items.Value = "Urunler";
                    Range support = ecurrentSheet.Range["A" + (erowCount + 1)];
                    support.EntireRow.Font.Bold = true;
                    support.Value = "Urunlerin Destek Degerleri";
		   // Coolumn sınıfından bir nesne ayağa kaldırılarak ürünlerin yatayda kapladıkları alana göre bulunuyor(kolon isimleri).
                    // Bu isimler bir dizide sakalanıyor.
                    Columns columns = new Columns();
                    string[] stringColumn = columns.writeColumn(columnCount);
                    for (int i = 1; i < columnCount - 1; i++)
                    {
                        Range valItem = currentSheet.Range["" + stringColumn[i] + 2];
                        valItem.EntireColumn.AutoFit();
                        var item = valItem.Value;
                        Range supp1 = ecurrentSheet.Range["" + stringColumn[i] + (erowCount)];
                        supp1.Value = item;
                        Range val = currentSheet.Range["" + stringColumn[i] + (rowCount2)];
                        var a = val.Value;
                        double div = a / total;
                        Range supp = ecurrentSheet.Range["" + stringColumn[i] + (erowCount + 1)];
                        supp.Value = (div);
                    }

                    // NEW ITEMSET (SUPPORT)
           	    // Destek değerine (supportValue) göre yeni bir veri seti oluşturuluyor.
                    // Bu destek değerini sağlamayan ürünler yeni veri setinde ihmal ediliyor.
                    int erowCount2 = ecurrentSheet.UsedRange.Rows.Count;
                    int arraysize = 0;
                    for (int i = 1; i < columnCount - 1; i++)
                    {
                        Range supp = ecurrentSheet.Range["" + stringColumn[i] + (erowCount2)];
                        double supV = supp.Value;
                        if (supV >= supportValue)
                        {
                            arraysize++;
                        }
                    }
                    string[] supString = new string[arraysize];
                    string[] supName = new string[arraysize];
                    int elementCount3 = 0;
                    for (int i = 1; i < columnCount - 1; i++)
                    {
                        Range supp = ecurrentSheet.Range["" + stringColumn[i] + (erowCount2)];
                        double supV = supp.Value;
                        if (supV >= supportValue)
                        {
                            Range supC = currentSheet.Range["" + stringColumn[i] + 2];
                            supString[elementCount3] = stringColumn[i];
                            elementCount3++;
                        }
                    }
	
		    // Eski veri setinden destek değerine uygun olan veriler yazdırılıyor(Yeni veri setinin elemanları).
                    Range val2 = ecurrentSheet.Range["A" + (erowCount2 + 2)];
                    val2.EntireRow.Font.Bold = true;
                    val2.Value = "Destek Degerine Uygun Urunler ";
                    val2.EntireColumn.AutoFit();
                    int each = 0;
                    for (char i = 'B'; i < 'B' + supString.Length; i++)
                    {
                        Range ival = currentSheet.Range["" + supString[each] + 2];
                        string rval = ival.Value;
                        supName[each] = rval;
                        Range val = ecurrentSheet.Range["" + i + (erowCount2 + 2)];
                        val.Value = rval;
                        each++;
                    }

                    // ASSOCIATION RULE
		    // Destek değerlerine uygun olan ürünleri ikili olarak ilişkileri analiz ediliyor.
                    // Combination sınıfından bir nesne ayağa kaldırılarak ilişki sayısı hesaplanıyor.(ürün adedinin ikili kombinasyonu)
                    Combination c1 = new Combination();
                    int sayacsize = Convert.ToInt32(c1.calcCombination(arraysize));
                    int[] sayacS = new int[sayacsize];
                    int element = 0;
                    bool flag = false;
		   // Pivot tablo içerisinde destek değerine uygun ürünler ikili olarak ele alınıp tüm alışverişler boyunca gezilerek kontrol ediliyor.
                    for (int k = 0; k < supString.Length; k++)
                    {
                        for (int j = 0; j < supString.Length; j++)
                        {
                            int sayac = 0;
                            flag = false;
                            for (int i = 0; i < total; i++)
                            {
                                if (k + j + 1 >= supString.Length)
                                {
                                    flag = true;
                                    break;
                                }
                                else
                                {
                                    Range ar = currentSheet.Range["" + supString[k] + (i + 3)];
                                    Range ar2 = currentSheet.Range["" + supString[k + (j + 1)] + (i + 3)];
				    // Pivot tablo içerinde bulunan bir alışverişte destek değerine uygun iki ürünün birlikte alınıp alınmadığı kontrol ediliyor.
                                    if (ar.Value != null && ar2.Value != null)
                                    {
                                        sayac++;
                                    }
                                }
                            }
                            if (flag != true)
                            {
                                sayacS[element] = sayac;
                                element++;
                            }
                        }
                    }

                    //SUPPORT VALUE NAME,SAYAC
		    // Destek değerine uygun olan ve ikili olarak değerlendirmeye alınan ürünlerin isimleri ve veri seti içerisindeki sıklıkları yazdırılıyor. 
                    Range val3 = ecurrentSheet.Range["A" + (erowCount2 + 4)];
                    val3.EntireRow.Font.Bold = true;
                    val3.Value = "Ikili Iliskiler ";
                    val3.EntireColumn.AutoFit();
                    int each2 = 0, say = 0;
                    for (char i = 'B'; i < 'B' + supString.Length - 1; i++)
                    {
                        for (int j = 0; j < supString.Length; j++)
                        {
                            if (each2 + j + 1 >= supString.Length)
                            {
                                break;
                            }
                            else
                            {
                                Range val = ecurrentSheet.Range["" + i + (erowCount2 + 4 + j)];
                                val.Value = "{" + supName[each2] + "," + supName[each2 + j + 1] + "}" + " = " + sayacS[say];
                                val.EntireColumn.AutoFit();
                                val.EntireRow.Font.Bold = true;
                            }
                            say++;
                        }
                        each2++;
                    }

                    //ASSOCIATION SUPPORT VALUE
	 	    // ikili olarak analiz edilen ürünlerin birlikteliklerinin destek değerleri hesaplanıyor ve yazdırılıyor
                    int rowcount = ecurrentSheet.UsedRange.Rows.Count;
                    Range val4 = ecurrentSheet.Range["A" + (rowcount + 2)];
                    val4.EntireRow.Font.Bold = true;
                    val4.Value = "Ikili Iliskilerin Destek Degerleri ";
                    val4.EntireColumn.AutoFit();
                    int each3 = 0, say2 = 0;

                    for (char i = 'B'; i < 'B' + supString.Length - 1; i++)
                    {
                        for (int j = 0; j < supString.Length; j++)
                        {
                            if (each3 + j + 1 >= supString.Length)
                            {
                                break;
                            }
                            else
                            {
                                if (Convert.ToDouble(sayacS[say2]) / total < supportValue)
                                {
				  // Birliktelikler destek değerinin altında kalıyorsa ekrana bilgi mesajı veriliyor.
                                    Range val = ecurrentSheet.Range["" + i + (rowcount + 2 + j)];
                                    val.Value = "{" + supName[each3] + "," + supName[each3 + j + 1] + "}" + " = " + "Bu iliski destek degerine uygun degildir";
                                    val.EntireColumn.AutoFit();
                                    val.EntireRow.Font.Bold = true;
                                }
                                else
                                {
				    // Birliktelikler destek değerinin altında kalmıyorsa ekrana hesaplana değer yazılıyor.
                                    Range val = ecurrentSheet.Range["" + i + (rowcount + 2 + j)];
                                    val.Value = "{" + supName[each3] + "," + supName[each3 + j + 1] + "}" + " = " + (Convert.ToDouble(sayacS[say2]) / total);
                                    val.EntireColumn.AutoFit();
                                    val.EntireRow.Font.Bold = true;
                                }
                            }
                            say2++;
                        }
                        each3++;
                    }
                    //RESULT VALUE ADD IN NEW WORKSHEET
                    //CREATE NEW WORKSHEET
		    // Sonuçların kullancıya bir tablo halinde yazdırılacağı yeni bir Excel çalışma sayfası açılıyor
                    Worksheet rcurrentSheet = Globals.ThisAddIn.GetActiveWorkSheet();
                    if (currentApp.Application.Sheets.Count < 5)
                    {
                        rcurrentSheet = (Excel.Worksheet)currentBook.Worksheets.Add();
                    }
                    else
                    {
                        rcurrentSheet = currentApp.Worksheets[5];
                    }
                    rcurrentSheet.Name = "Result";

                    //ASSOCIATION COMFIDENCE VALUE
		    // Birlikteliği olan ürünlerin güven değerleri hesaplanıyor ve yazdırılıyor
                    int saysize2 = 0;
                    for (int i = 0; i < sayacS.Length; i++)
                    {
                        if ((Convert.ToDouble(sayacS[i]) / total) > supportValue)
                        {
                            saysize2++;
                        }

                    }
		    // Birlikteliğe giren ürünlerin isimleri ve güven değerlerini bir daha sonra işlenmek üzere bir dizide tuluyor.
                    string[] comfValS = new string[saysize2];
                    double[] comfVal = new double[saysize2];
                    int rowcount2 = ecurrentSheet.UsedRange.Rows.Count;
                    Range val5 = ecurrentSheet.Range["A" + (rowcount2 + 2)];
                    val5.EntireRow.Font.Bold = true;
                    val5.Value = "Ikili Iliskilerin Guven Degerleri ";
                    val5.EntireColumn.AutoFit();
                    int each4 = 0, say3 = 0, sayrc = 0;
                    double[] SupStringSup = new double[supString.Length];
                    for (int k = 0; k < supString.Length; k++)
                    {
                        Range val = ecurrentSheet.Range["" + supString[k] + (erowCount + 1)];
                        SupStringSup[k] = Convert.ToDouble(val.Value);
                    }

                    for (char i = 'B'; i < 'B' + supString.Length - 1; i++)
                    {
                        for (int j = 0; j < supString.Length; j++)
                        {
                            if (each4 + j + 1 >= supString.Length)
                            {
                                break;
                            }
                            else
                            {
                                if (Convert.ToDouble(sayacS[say3]) / total < supportValue)
                                {
				    // Birliktelikler destek değerinin altında kalıyorsa ekrana bilgi mesajı veriliyor.
                                    Range val = ecurrentSheet.Range["" + i + (rowcount2 + 2 + j)];
                                    val.Value = "{" + supName[each4] + "," + supName[each4 + j + 1] + "}" + " = " + "Bu iliski destek degerine uygun degildir";
                                    val.EntireColumn.AutoFit();
                                    val.EntireRow.Font.Bold = true;
                                }
                                else
                                {
                                     // Birliktelikler destek değerinin altında kalmıyorsa ekrana hesaplana değer yazılıyor.
                                    Range val = ecurrentSheet.Range["" + i + (rowcount2 + 2 + j)];
                                    val.Value = "{" + supName[each4] + "," + supName[each4 + j + 1] + "}" + " = " + ((Convert.ToDouble(sayacS[say3]) / total) / SupStringSup[each4]);
                                    val.EntireColumn.AutoFit();
                                    val.EntireRow.Font.Bold = true;
                                    comfValS[sayrc] = "{" + supName[each4] + "," + supName[each4 + j + 1] + "}" + " = ";
                                    comfVal[sayrc] = (Convert.ToDouble(sayacS[say3]) / total) / SupStringSup[each4];
                                    sayrc++;
                                }
                            }
                            say3++;
                        }
                        each4++;
                    }

                    //ASSOCIATION RECOMFİDENCE VALUE    
	            // Birlikteliği olan ürünlerin ters ilişki şeklinde güven değerleri hesaplanıyor ve yazdırılıyor
                    int saysize1 = 0;
                    for (int i = 0; i < sayacS.Length; i++)
                    {
                        if ((Convert.ToDouble(sayacS[i]) / total) > supportValue)
                        {
                            saysize1++;
                        }
                    }
	            // Birlikteliğe giren ürünlerin isimleri ve ters ilişki güven değerlerini bir daha sonra işlenmek üzere bir dizide tuluyor.
                    string[] ncomfValS = new string[saysize1];
                    double[] ncomfVal = new double[saysize1];
                    int rowcount3 = ecurrentSheet.UsedRange.Rows.Count;
                    Range val6 = ecurrentSheet.Range["A" + (rowcount3 + 2)];
                    val6.EntireRow.Font.Bold = true;
                    val6.Value = "Ters Ikili Iliskilerin Guven Degerleri ";
                    val6.EntireColumn.AutoFit();
                    int each5 = 0, say4 = 0, sayrnc = 0;
                    for (char i = 'B'; i < 'B' + supString.Length - 1; i++)
                    {
                        for (int j = 0; j < supString.Length; j++)
                        {
                            if (each5 + j + 1 >= supString.Length)
                            {
                                break;
                            }
                            else
                            {
                                if ((Convert.ToDouble(sayacS[say4]) / total) < supportValue)
                                {
				    // Birliktelikler destek değerinin altında kalıyorsa ekrana bilgi mesajı veriliyor.
                                    Range val = ecurrentSheet.Range["" + i + (rowcount3 + 2 + j)];
                                    val.Value = "{" + supName[each5 + j + 1] + "," + supName[each5] + "}" + " = " + "Bu iliski destek degerine uygun degildir";
                                    val.EntireColumn.AutoFit();
                                    val.EntireRow.Font.Bold = true;
                                }
                                else
                                {
				    // Birliktelikler destek değerinin altında kalmıyorsa ekrana hesaplana değer yazılıyor.
                                    Range val = ecurrentSheet.Range["" + i + (rowcount3 + 2 + j)];
                                    val.Value = "{" + supName[each5 + j + 1] + "," + supName[each5] + "}" + " = " + ((Convert.ToDouble(sayacS[say4]) / total) / SupStringSup[each5 + j + 1]);
                                    val.EntireColumn.AutoFit();
                                    val.EntireRow.Font.Bold = true;
                                    ncomfValS[sayrnc] = "{" + supName[each5 + j + 1] + "," + supName[each5] + "}" + " = ";
                                    ncomfVal[sayrnc] = (Convert.ToDouble(sayacS[say4]) / total) / SupStringSup[each5 + j + 1];
                                    sayrnc++;
                                }
                            }
                            say4++;
                        }
                        each5++;
                    }

                    //ASSOCIATION LIFT VALUE
	            // Birlikteliği olan ürünlerin kaldıraç(lift) değerleri hesaplanıyor ve yazdırılıyor
                    int saysize = 0;
                    for (int i = 0; i < sayacS.Length; i++)
                    {
                        if ((Convert.ToDouble(sayacS[i]) / total) > supportValue)
                        {
                            saysize++;
                        }
                    }
		    // Birlikteliğe giren ürünlerin isimleri ve lift değerlerini bir daha sonra işlenmek üzere bir dizide tuluyor.
                    string[] liftValS = new string[saysize];
                    double[] liftVal = new double[saysize];
                    int rowcount4 = ecurrentSheet.UsedRange.Rows.Count;
                    Range val7 = ecurrentSheet.Range["A" + (rowcount4 + 2)];
                    val7.EntireRow.Font.Bold = true;
                    val7.Value = "Ikili Iliskilerin Lift Degerleri ";
                    val7.EntireColumn.AutoFit();
                    int each6 = 0, say5 = 0, sayr = 0;
                    for (char i = 'B'; i < 'B' + supString.Length - 1; i++)
                    {
                        for (int j = 0; j < supString.Length; j++)
                        {
                            if (each6 + j + 1 >= supString.Length)
                            {
                                break;
                            }
                            else
                            {
                                if ((Convert.ToDouble(sayacS[say5]) / total) < supportValue)
                                {
				    // Birliktelikler destek değerinin altında kalıyorsa ekrana bilgi mesajı veriliyor.
                                    Range val = ecurrentSheet.Range["" + i + (rowcount4 + 2 + j)];
                                    val.Value = "{" + supName[each6] + "," + supName[each6 + j + 1] + "}" + " = " + "Bu iliski destek degerine uygun degildir";
                                    val.EntireColumn.AutoFit();
                                    val.EntireRow.Font.Bold = true;
                                }
                                else
                                {
				    // Birliktelikler destek değerinin altında kalmıyorsa ekrana hesaplana değer yazılıyor.
                                    Range val = ecurrentSheet.Range["" + i + (rowcount4 + 2 + j)];
                                    val.Value = "{" + supName[each6] + "," + supName[each6 + j + 1] + "}" + " = " + ((Convert.ToDouble(sayacS[say5]) / total) / (SupStringSup[each6] * SupStringSup[each6 + j + 1]));
                                    val.EntireColumn.AutoFit();
                                    val.EntireRow.Font.Bold = true;
                                    liftValS[sayr] = "{" + supName[each6] + "," + supName[each6 + j + 1] + "}" + " = ";
                                    liftVal[sayr] = (Convert.ToDouble(sayacS[say5]) / total) / (SupStringSup[each6] * SupStringSup[each6 + j + 1]);
                                    sayr++;

                                }
                            }
                            say5++;
                        }
                        each6++;
                    }

                    //PRINT RESULT VALUE
		    // Destek değerine uygun olan birlikteliklerin güven değerleri yazdırılıyor.
                    Range title = rcurrentSheet.Range["A1"];
                    title.Value = "Comfidence";
                    title.EntireColumn.AutoFit();
                    title.EntireRow.Font.Bold = true;
                    rcurrentSheet.get_Range("A1", "C1").VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    rcurrentSheet.get_Range("A1", "C1").HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    Swap swp = new Swap();
                    swp.swap(comfVal, comfValS);
                    ResultPrint resPrnt = new ResultPrint();
                    rcurrentSheet = resPrnt.printResult(rcurrentSheet, comfVal, comfValS, "A");
			
		    // Destek değerine uygun olan birlikteliklerin ters ilişki güven değerleri yazdırılıyor.
                    Range title2 = rcurrentSheet.Range["B1"];
                    title2.Value = "RComfidence";
                    title2.EntireColumn.AutoFit();
                    title2.EntireRow.Font.Bold = true;
                    swp.swap(ncomfVal, ncomfValS);
                    rcurrentSheet = resPrnt.printResult(rcurrentSheet, ncomfVal, ncomfValS, "B");

		    // Destek değerine uygun olan birlikteliklerin lift değerleri yazdırılıyor.
                    Range title3 = rcurrentSheet.Range["C1"];
                    title3.Value = "Lift";
                    title3.EntireColumn.AutoFit();
                    title3.EntireRow.Font.Bold = true;
                    swp.swap(liftVal, liftValS);
                    rcurrentSheet = resPrnt.printResult(rcurrentSheet, liftVal, liftValS, "C");

                    //CREATE TABLE
		    // birliktelik sonuçları tablo haline getiriliyor.
                    cPT.createTable(rcurrentSheet);
                }
		// Kullanıcı yanlış birşeyler yaparsa hata mesajı ekrana basılıyor.
                catch
                {
                    MessageBox.Show("Birşeyler ters gitti, uygulama adımlarını kontrol ediniz !");
                }
            }
	    // if koşulu sağlanmazsa kullanıcıya verilecek mesaj ekrana basılıyor.
            else MessageBox.Show("Bir Destek Değeri Belirleyiniz !");
        }

	// Hakkında butonunun eventi.
        private void btn_Info_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("Veri setinizi; 'Tarih', 'Saat', 'FişNo', 'ÜrünAdı' şeklinde düzenledikten sonra A2 satırından başlayacak şekilde aktarınız!"
                + " Ayrıca destek değerini(Support Value) ondalıklı bir sayı olarak belirleyin.\n\n\n"
                + "Bu uygulama, Pamukkale Üniversitesi Yönetim Bilişim Sistemleri Bölümü öğrencisi Halil KAYACI'nın bitirme projesidir.\n\n"
                + "Soru ve görüşleriniz için : halilkayaci@gmail.com","Hakkında", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

	 // Yardım butonunun eventi.
        private void btn_Help_Click(object sender, RibbonControlEventArgs e)
        {
            Help help = new Help();
            help.Show();
        }
    }
}
