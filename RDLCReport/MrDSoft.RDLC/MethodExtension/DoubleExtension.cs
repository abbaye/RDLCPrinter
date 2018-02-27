using System;

namespace DSoft.MethodExtension
{

    /// <remarks>
    /// CREDIT : 2013 Derek Tremblay (abbaye)
    /// https://github.com/abbaye/RDLCPrinter
    /// </remarks>
    public static class DoubleExtension
    {
        public static double Round(this double s, int digit)
        {
            return Math.Round(s, digit);
        }

    }
}
