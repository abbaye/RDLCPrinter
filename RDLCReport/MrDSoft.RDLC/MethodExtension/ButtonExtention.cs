using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace DSoft.MethodExtension
{
    public static class ButtonExtention
    {
        /// <summary>
        /// Desactive un boutton et baisse son opacité a 50%
        /// </summary>
        /// <param name="button"></param>
        public static void DisableButton(this Button button)
        {
            button.Opacity = 0.5;
            button.IsEnabled = false;
        }

        /// <summary>
        /// Active un boutton et ajuste son opacité a 100%
        /// </summary>
        /// <param name="button"></param>
        public static void EnableButton(this Button button)
        {
            button.Opacity = 1;
            button.IsEnabled = true;
        }


        public static void PerformClick(this System.Windows.Controls.Button btn)
        {
            btn.RaiseEvent(new RoutedEventArgs(System.Windows.Controls.Button.ClickEvent));
        }
        
    }
}
