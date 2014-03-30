using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace DSoft.RDLCReport
{
    /// <summary>
    /// Logique d'interaction pour LightIntergerSpinner.xaml
    /// </summary>
    public partial class LightIntergerSpinner : UserControl
    {
        private int? _minimum = null;
        private int? _maximum = null;


        public LightIntergerSpinner()
        {
            InitializeComponent();
        }

        public int? Value
        {
            get
            {
                try
                {
                    return Convert.ToInt32(NumPager.Text);
                }
                catch
                {
                    return null;
                }
            }
            set
            {                
                NumPager.Text = value.ToString();
            }
        }

        public int? Minimum
        {
            get
            {
                return this._minimum;
            }
            set
            {
                this._minimum = value;                                
            }
        }

        public int? Maximum
        {
            get
            {
                return this._maximum;
            }
            set
            {
                this._maximum = value;

            }
        }

        private void SpinnerUp_Click(object sender, RoutedEventArgs e)
        {
            this.Value += 1;            
        }

        private void SpinnerDown_Click(object sender, RoutedEventArgs e)
        {
            this.Value -= 1;
        }
    }
}
