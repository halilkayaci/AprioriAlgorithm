using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AprioriAlgorithm
{
    class Combination
    {
	// Parametre olarak verilen sayının faktöriyelini hesaplayan metod.
        public double calcFactorial(int number)
        {
            double factorial = 1;
            for (int i = 1; i <= number; i++)
            {
                factorial *= i;
            }
            return factorial;
        }
	// Parametre olarak verilen sayının 2'li kombinasyonunu hesaplayan metod.
        public double calcCombination(int number)
        {
            double sonuc = 0;
            sonuc = calcFactorial(number) / ((calcFactorial(2) * calcFactorial(number - 2)));
            return sonuc;
        }
    }
}
