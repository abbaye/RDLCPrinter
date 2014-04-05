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

using DSoft.MethodExtension;

namespace DSoft.RDLCReport
{
    /// <summary>
    /// LightIntergerSpinner
    /// <remarks>
    /// CREDIT : 2013-2014 Derek Tremblay (abbaye)
    /// https://rdlcprinter.codeplex.com/
    /// </remarks>
    /// </summary>
    public partial class LightIntergerSpinner : UserControl
    {
        private int _minimum = 0;
        private int _maximum = 50;
        private bool _isShowSpinnerButton = true;

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

                CheckRange();
                UpdateButton();

                if (ValueChanged != null)
                    ValueChanged(this, new EventArgs());
            }
        }

        public bool IsShowSpinnerButton
        {
            get
            {
                return this._isShowSpinnerButton;
            }
            set
            {
                this._isShowSpinnerButton = value;

                if (value)
                    ButtonColumn.Width = new GridLength(20);
                else
                    ButtonColumn.Width = new GridLength(0);
            }
        }

        private void UpdateButton()
        {
            SpinnerUp.DisableButton();
            SpinnerDown.DisableButton();

            if (this.Value > this._minimum)
                SpinnerDown.EnableButton();

            if (this.Value < this._maximum)
                SpinnerUp.EnableButton();
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
                UpdateButton();
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
                UpdateButton();
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

            CheckRange();
            UpdateButton();

            if (UpButtonClick != null)
                UpButtonClick(this, new EventArgs());
        }

        private void SpinnerDown_Click(object sender, RoutedEventArgs e)
        {
            if (this.Value > this._minimum)
                this.Value -= 1;

            CheckRange();
            UpdateButton();

            if (DownButtonClick != null)
                DownButtonClick(this, new EventArgs());
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            CheckRange();
            UpdateButton();
        }
    }
}
