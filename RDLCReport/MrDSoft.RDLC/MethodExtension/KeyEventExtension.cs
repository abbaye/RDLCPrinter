using System.Windows.Input;

namespace DSoft.MethodExtension
{
    /// <remarks>
    /// CREDIT : 2013-2014 Derek Tremblay (abbaye)
    /// https://github.com/abbaye/RDLCPrinter
    /// </remarks>
    public static class KeyEventExtension
    {
        public static bool IsNumericKey(this KeyEventArgs e)
        {
            if ((e.Key >= Key.D0 && e.Key <= Key.D9) || (e.Key >= Key.NumPad0 && e.Key <= Key.NumPad9))
                return true;
            else
                return false;
        }
    }
}
