using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DSoft.MethodExtension
{
    public static class DoubleExtension
    {
        public static double Round(this double s, int digit)
        {
            return Math.Round(s, digit);
        }

    }
}
