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
        private int _minimum = 0;
        private int _maximum = 50;

        public event EventHandler ValueChanged;
        public event EventHandler UpButtonClick;
        public event EventHandler DownButtonClick;
        private bool _IsShowCurrentToMaximumValue;


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
                    if (this._IsShowCurrentToMaximumValue)                    
                        return Convert.ToInt32(NumPager.Text.Split('/')[0].Trim());                    
                    else                    
                        return Convert.ToInt32(NumPager.Text);                    
                }
                catch
                {
                    return null;
                }
            }
            set
            {
                if (this._IsShowCurrentToMaximumValue)
                {
                    NumPager.Text = value.ToString() + " / " + this._maximum.ToString();
                }else
                    NumPager.Text = value.ToString();

                if (ValueChanged != null)
                    ValueChanged(this, new EventArgs());
            }
        }

        public int Minimum
        {
            get
            {
                return this._minimum;
            }
            set
            {
                this._minimum = value;

                CheckRange();
            }
        }

        public int Maximum
        {
            get
            {
                return this._maximum;
            }
            set
            {
                this._maximum = value;

                CheckRange();
            }
        }

        public bool IsShowCurrentToMaximumValue
        {
            get
            {
                return this._IsShowCurrentToMaximumValue;
            }
            set
            {
                this._IsShowCurrentToMaximumValue = value;

                this.Value = this.Value;
            }
        }
        

        private void CheckRange()
        {
            if (this.Value > this._maximum)
            {
                this.Value = this._maximum;

                if (ValueChanged != null)
                    ValueChanged(this, new EventArgs());
            }

            if (this.Value < this._minimum)
            {
                this.Value = this._minimum;

                if (ValueChanged != null)
                    ValueChanged(this, new EventArgs());
            }
        }

        private void SpinnerUp_Click(object sender, RoutedEventArgs e)
        {
            if (this.Value < this._maximum)
                this.Value += 1;

            if (UpButtonClick != null)
                UpButtonClick(this, new EventArgs());
        }

        private void SpinnerDown_Click(object sender, RoutedEventArgs e)
        {
            if (this.Value > this._minimum)
                this.Value -= 1;

            if (DownButtonClick != null)
                DownButtonClick(this, new EventArgs());
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            CheckRange();
        }
    }
}
