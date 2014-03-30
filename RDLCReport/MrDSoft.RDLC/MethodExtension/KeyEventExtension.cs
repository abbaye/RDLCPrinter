using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace DSoft.MethodExtension
{
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
