using System;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using DSoft.MethodExtension;

namespace DSoft.RDLC
{
    /// <summary>
    /// RDLCSlider
    /// <remarks>
    /// CREDIT : 2013-2018 Derek Tremblay (abbaye), 2013 Martin Savard
    /// https://github.com/abbaye/RDLCPrinter
    /// </remarks>
    /// </summary>
    public partial class LightDoubleSlider : UserControl
    {
        private double _sliderMinimum;
        private double _sliderMaximum = 1;
        private double _sliderSmallChange = 1;
        private int _sliderInterval = 1;
        private double _sliderTickFrequency = 1;
        private bool _slideIsSnapToTickEnabled = true;
        private double _sliderValue;
        public event EventHandler FirstButtonClick;
        public event EventHandler PrevioustButtonClick;
        public event EventHandler NextButtonClick;
        public event EventHandler LastButtonClick;
        public event EventHandler ValueChanged;



        public LightDoubleSlider()
        {
            InitializeComponent();
        }


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
                return _sliderMaximum;
            }

            set
            {
                _sliderMaximum = value;
                ChartSlider.Maximum = _sliderMaximum;

            }
        }


        /// <summary>
        /// Get or set the Smallchange 
        /// </summary>
        public double SmallChange
        {
            get
            {
                return _sliderSmallChange;
            }

            set
            {
                _sliderSmallChange = value;
                ChartSlider.SmallChange = _sliderSmallChange;
                ChartSlider.LargeChange = _sliderSmallChange;
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
                return _sliderInterval;
            }

            set
            {
                _sliderInterval = value;
                ChartSlider.Interval = _sliderInterval;

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
                return _sliderTickFrequency;
            }

            set
            {
                _sliderTickFrequency = value;
                ChartSlider.TickFrequency = _sliderTickFrequency;

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
                return _slideIsSnapToTickEnabled;
            }

            set
            {
                _slideIsSnapToTickEnabled = value;
                ChartSlider.IsSnapToTickEnabled = _slideIsSnapToTickEnabled;

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
                ChartSlider.Value = _sliderValue;

                ValueChanged?.Invoke(this, new EventArgs());
            }
        }
        #endregion //Properties

        #region Events

        private void FirstButton_Click(object sender, RoutedEventArgs e)
        {
            UpdateButton();

            FirstButtonClick?.Invoke(this, new EventArgs());
        }

        private void PreviousButton_Click(object sender, RoutedEventArgs e)
        {
            UpdateButton();

            PrevioustButtonClick?.Invoke(this, new EventArgs());
        }

        private void NextButton_Click(object sender, RoutedEventArgs e)
        {
            UpdateButton();

            NextButtonClick?.Invoke(this, new EventArgs());
        }

        private void LastButton_Click(object sender, RoutedEventArgs e)
        {
            UpdateButton();

            LastButtonClick?.Invoke(this, new EventArgs());
        }

        private void ChartSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            _sliderValue = ChartSlider.Value;

            UpdateButton();

            ValueChanged?.Invoke(this, new EventArgs());
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
