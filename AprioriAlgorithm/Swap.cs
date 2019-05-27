using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AprioriAlgorithm
{
    class Swap
    {
        // Parametre olarak verilen double diziyi büyükten küçüye doğru sıralarken
        // aynı zamanda değerlerin karşılık geldiği string dizisini de düzenleyen metod.
        public void swap(double[] swapArray, string[] swapArrayS)
        {
            double swap; string swapS;
            for (int i = 0; i < swapArray.Length; i++)
            {
                for (int j = 0; j < swapArray.Length; j++)
                {
                    if (swapArray[i] > swapArray[j])
                    {
                        swap = swapArray[i];
                        swapArray[i] = swapArray[j];
                        swapArray[j] = swap;

                        swapS = swapArrayS[i];
                        swapArrayS[i] = swapArrayS[j];
                        swapArrayS[j] = swapS;
                    }

                }
            }
        }
    }
}
