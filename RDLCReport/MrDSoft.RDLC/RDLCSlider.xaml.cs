using System;
using System.Collections.Generic;
using System.ComponentModel;
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

namespace DSoft.RDLC
{
    /// <summary>
    /// RDLCSlider
    /// <remarks>
    /// CREDIT : 2013-2014 Derek Tremblay (abbaye), Martin Savard
    /// https://rdlcprinter.codeplex.com/
    /// </remarks>
    /// </summary>
    public partial class RDLCSlider : UserControl
    {

        private double _sliderMinimum = 0;
        private double _sliderMaximum = 1;
        private double _sliderSmallChange = 1;
        private int _sliderInterval = 1;
        private double _sliderTickFrequency = 1;
        private bool _slideIsSnapToTickEnabled = true;
        private double _sliderValue = 0;
        public event EventHandler FirstButtonClick;
        public event EventHandler PrevioustButtonClick;
        public event EventHandler NextButtonClick;
        public event EventHandler LastButtonClick;
        public event EventHandler ValueChanged;


        #region Properties
        /// <summary>
        /// Get ou set the minimum value
        /// </summary>
        public double Minimum
        {
            get
            {
                return ChartSlider.Minimum = _sliderMinimum;
            }
            set
            {
                _sliderMinimum = value;
                ChartSlider.Minimum = _sliderMinimum;
            }
        }


        /// <summary>
        /// Get or set the maximum value
        /// </summary>
        public double Maximum
        {
            get
            {
                return this._sliderMaximum;
            }

            set
            {
                this._sliderMaximum = value;
                this.ChartSlider.Maximum = this._sliderMaximum;

            }
        }


        /// <summary>
        /// Get or set the Smallchange 
        /// </summary>
        public double SmallChange
        {
            get
            {
                return this._sliderSmallChange;
            }

            set
            {
                this._sliderSmallChange = value;
                this.ChartSlider.SmallChange = this._sliderSmallChange;
                this.ChartSlider.LargeChange = this._sliderSmallChange;
            }
        }


        /// <summary>
        /// Get or set the interval
        /// </summary>
        [DefaultValue(1)]
        public int Interval
        {
            get
            {
                return this._sliderInterval;
            }

            set
            {
                this._sliderInterval = value;
                this.ChartSlider.Interval = this._sliderInterval;

            }
        }


        /// <summary>
        /// Get or set the TickFrequency
        /// </summary>
        [DefaultValue(1)]
        public double TickFrequency
        {
            get
            {
                return this._sliderTickFrequency;
            }

            set
            {
                this._sliderTickFrequency = value;
                this.ChartSlider.TickFrequency = this._sliderTickFrequency;

            }
        }


        /// <summary>
        /// Get or set the IsSnapToTickEnabled
        /// </summary>
        [DefaultValue(true)]
        public bool IsSnapToTickEnabled
        {
            get
            {
                return this._slideIsSnapToTickEnabled;
            }

            set
            {
                this._slideIsSnapToTickEnabled = value;
                this.ChartSlider.IsSnapToTickEnabled = this._slideIsSnapToTickEnabled;

            }
        }


        /// <summary>
        /// Get or set the curent value of slider
        /// </summary>
        [DefaultValue(0)]
        public double Value
        {
            get
            {
                return _sliderValue;
            }
            set
            {
                _sliderValue = value;
                this.ChartSlider.Value = this._sliderValue;

                if (ValueChanged != null)
                    ValueChanged(this, new EventArgs());
            }
        }
        #endregion //Properties


        public RDLCSlider()
        {
            InitializeComponent();
        }


        #region Events

        private void FirstButton_Click(object sender, RoutedEventArgs e)
        {
            
            if (FirstButtonClick != null)
                FirstButtonClick(this, new EventArgs());
        }

        private void PreviousButton_Click(object sender, RoutedEventArgs e)
        {
            if (PrevioustButtonClick != null)
                PrevioustButtonClick(this, new EventArgs());
        }

        private void NextButton_Click(object sender, RoutedEventArgs e)
        {
            if (NextButtonClick != null)
                NextButtonClick(this, new EventArgs());
        }

        private void LastButton_Click(object sender, RoutedEventArgs e)
        {
            if (LastButtonClick != null)
                LastButtonClick(this, new EventArgs());
        }

        private void ChartSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            _sliderValue = ChartSlider.Value;

            if (ValueChanged != null)
                ValueChanged(this, new EventArgs());
        }

        #endregion //Events.


        #region Methode

        public void UpdateButton()
        {
            if (ChartSlider.Value == 0)
            {
                FirstButton.DisableButton();
                PreviousButton.DisableButton();
            }
            else
            {
                FirstButton.EnableButton();
                PreviousButton.EnableButton();
            }

            if (ChartSlider.Value == ChartSlider.Maximum)
            {
                NextButton.DisableButton();
                LastButton.DisableButton();
            }
            else
            {
                NextButton.EnableButton();
                LastButton.EnableButton(); 
            }
        }
        #endregion //Methode
    }
}
