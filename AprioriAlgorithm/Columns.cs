using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AprioriAlgorithm
{
    class Columns
    {
        // Parametre olarak verilen sütun sayısı kadar ürünlerin excel sütunlarına yerleşmesini sağlamak amacıyla excel sütün isimlerini
        // oluşturan metod.
        public string[] writeColumn(int columnCount)
        {
            var modCloumn = columnCount % 26;
            int resColumn = columnCount / 26;
            string[] character = {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K",
                                            "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V",
                                            "W", "X", "Y", "Z" };
            string[] chColumn = new string[columnCount];
            var element = 0; var control = 0;
            if (resColumn != 0)
            {
                for (int i = 0; i < resColumn; i++)
                {
                    for (int j = 0; j < character.Length; j++)
                    {
                        if (i == 0)
                        {
                            chColumn[element] = character[j];
                            element++;
                        }

                        else
                        {
                            chColumn[element] = character[control] + character[j];
                            element++;
                        }
                    }
                    if (i != 0) control++;
                }
                for (int i = 0; i < modCloumn; i++)
                {
                    chColumn[element] = character[resColumn - 1] + character[i];
                    element++;
                }
            }
            else
            {
                for (int i = 0; i < columnCount; i++)
                {
                    chColumn[i] = character[i];
                }
            }

            return chColumn;

        }

    }
}
