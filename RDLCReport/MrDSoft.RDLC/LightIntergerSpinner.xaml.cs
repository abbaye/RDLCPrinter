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
                    if (_IsShowCurrentToMaximumValue)                    
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
                if (_IsShowCurrentToMaximumValue)
                {
                    NumPager.Text = value.ToString() + " / " + _maximum.ToString();
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
                return _isShowSpinnerButton;
            }
            set
            {
                _isShowSpinnerButton = value;

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

            if (Value > _minimum)
                SpinnerDown.EnableButton();

            if (Value < _maximum)
                SpinnerUp.EnableButton();
        }

        public int Minimum
        {
            get
            {
                return _minimum;
            }
            set
            {
                _minimum = value;

                CheckRange();
                UpdateButton();
            }
        }

        public int Maximum
        {
            get
            {
                return _maximum;
            }
            set
            {
                _maximum = value;

                CheckRange();
                UpdateButton();
            }
        }

        public bool IsShowCurrentToMaximumValue
        {
            get
            {
                return _IsShowCurrentToMaximumValue;
            }
            set
            {
                _IsShowCurrentToMaximumValue = value;

                Value = Value;
            }
        }
        

        private void CheckRange()
        { 

            if (Value > _maximum)
            {
                Value = _maximum;

                if (ValueChanged != null)
                    ValueChanged(this, new EventArgs());
            }

            if (Value < _minimum)
            {
                Value = _minimum;

                if (ValueChanged != null)
                    ValueChanged(this, new EventArgs());
            }
        }

        private void SpinnerUp_Click(object sender, RoutedEventArgs e)
        {
            if (Value < _maximum)
                Value += 1;

            CheckRange();
            UpdateButton();

            if (UpButtonClick != null)
                UpButtonClick(this, new EventArgs());
        }

        private void SpinnerDown_Click(object sender, RoutedEventArgs e)
        {
            if (Value > _minimum)
                Value -= 1;

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
