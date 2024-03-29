using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;

namespace TimePunch.Rdlc.MethodExtension
{
    /// <remarks>
    /// CREDIT : 2013-2018 Derek Tremblay (abbaye), 2013 Martin Savard
    /// https://github.com/abbaye/RDLCPrinter
    /// </remarks>
    public static class ButtonExtention
    {
        /// <summary>
        /// Desactive un boutton et baisse son opacité a 50%
        /// </summary>
        public static void DisableButton(this Button button)
        {
            button.Opacity = 0.5;
            button.IsEnabled = false;
        }

        /// <summary>
        /// Active un boutton et ajuste son opacité a 100%
        /// </summary>
        public static void EnableButton(this Button button)
        {
            button.Opacity = 1;
            button.IsEnabled = true;
        }
        
        public static void PerformClick(this Button btn) => 
            btn.RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
    }
}
