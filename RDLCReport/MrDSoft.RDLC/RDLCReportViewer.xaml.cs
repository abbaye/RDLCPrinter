//////////////////////////////////////////////
// MIT - 2012-2020
// Author : Derek Tremblay (derektremblay666@gmail.com)
// Contributor : Martin Savard (2013)
// https://github.com/abbaye/RDLCPrinter
//////////////////////////////////////////////

using System.Windows;
using System.Windows.Input;
using ModernWpf.Controls;
using TimePunch.Rdlc.MethodExtension;

namespace TimePunch.Rdlc
{
    /// <summary>
    /// RDLC Preview user control
    /// </summary>
    public partial class RdlcReportViewer
    {
        private RdlcPrinter _report;
        private int _pos;

        public RdlcReportViewer()
        {
            InitializeComponent();
            UpdateToolBarButton();
        }

        /// <summary>
        /// Call refresh button
        /// </summary>        
        private void CmdRefresh_Click(object sender, RoutedEventArgs e) => RefreshControl();

        /// <summary>
        /// Get or Set the RDLC report
        /// </summary>
        public RdlcPrinter Report
        {
            get
            {
                return _report;
            }
            set
            {
                _report = value;
                RefreshControl();
                GiveFocus();
            }
        }

        /// <summary>
        /// Give the focus to usercontrol
        /// </summary>
        public void GiveFocus()
        {
            FocusManager.SetFocusedElement(this, PreviewImage);
            Keyboard.Focus(PreviewImage);
        }

        /// <summary>
        /// Refresh logic.         
        /// </summary>
        public void RefreshControl()
        {
            if (_report == null) return;

            _report.Refresh();

            DispatcherHelper.DoEvents(); //Clear the ui message queud

            LoadImage();

            _pos = 1;
            PageSpinner.Maximum = _report.PagesCount;
            PageSpinner.Value = _pos;
            ChangeImage(_pos);
        }

        /// <summary>
        /// Update the toolbar button
        /// </summary>
        private void UpdateToolBarButton()
        {
            if (_pos < 0) 
                return;

            if (_report != null)
            {
                ButtonExtention.EnableButton(TBBRefresh);
                ExportMenu.IsEnabled = true;
                ExportMenu.Opacity = 1;
                ButtonExtention.EnableButton(TBBPrintWithProperties);
                ZoomInfoStackPanel.Visibility = Visibility.Visible;
                ZoomPopupButton.Visibility = Visibility.Visible;
                ZoomPopupButton.IsEnabled = true;
                ZoomPopupButton.Opacity = 1;
            }
            else
            {
                ButtonExtention.DisableButton(TBBRefresh);
                ExportMenu.IsEnabled = false;
                ExportMenu.Opacity = 0.5;
                ButtonExtention.DisableButton(TBBPrintWithProperties);
                ZoomInfoStackPanel.Visibility = Visibility.Collapsed;
                ZoomPopupButton.Visibility = Visibility.Collapsed;
                ZoomPopupButton.IsEnabled = false;
                ZoomPopupButton.Opacity = 0.5;
            }

            if (_pos == 1)
            {
                PagerSeparator.Visibility = Visibility.Visible;
                ButtonExtention.DisableButton(PreviousImage);
                ButtonExtention.DisableButton(FirstImage);
                ButtonExtention.EnableButton(NextImage);
                ButtonExtention.EnableButton(LastImage);
            }
            else
            {
                if (_report != null && _pos == _report.PagesCount)
                {
                    ButtonExtention.DisableButton(NextImage);
                    ButtonExtention.DisableButton(LastImage);
                    ButtonExtention.EnableButton(PreviousImage);
                    ButtonExtention.EnableButton(FirstImage);
                }
                else
                {
                    ButtonExtention.EnableButton(PreviousImage);
                    ButtonExtention.EnableButton(NextImage);
                    ButtonExtention.EnableButton(FirstImage);
                    ButtonExtention.EnableButton(LastImage);
                }
            }

            if (_report != null && _report.PagesCount > 1 && _pos>0)
            {
                PagerSeparator.Visibility = Visibility.Visible;
                PreviousImage.Visibility = Visibility.Visible;
                NextImage.Visibility = Visibility.Visible;
                FirstImage.Visibility = Visibility.Visible;
                LastImage.Visibility = Visibility.Visible;
                PageSpinner.Visibility = Visibility.Visible;
            }
            else
            {
                PagerSeparator.Visibility = Visibility.Collapsed;
                PreviousImage.Visibility = Visibility.Collapsed;
                NextImage.Visibility = Visibility.Collapsed;
                FirstImage.Visibility = Visibility.Collapsed;
                LastImage.Visibility = Visibility.Collapsed;
                PageSpinner.Visibility = Visibility.Collapsed;
            }
        }

        /// <summary>
        /// Load image for the preview
        /// </summary>
        private void LoadImage()
        {
            if (_pos == 0)
            {
                UpdateToolBarButton();
                PreviousImage.IsEnabled = false;
                PreviousImage.Opacity = 0.5;
            }
        }

        /// <summary>
        /// Chage page to the position
        /// </summary>
        private void ChangeImage(int position)
        {
            if (_report == null || Application.Current.MainWindow == null)
                return;

            try
            {
                Application.Current.MainWindow.Cursor = Cursors.Wait;

                var pagecount = _report.GetBitmapDecoder().Frames.Count;

                //Check interval
                if (position <= 0)
                    position = 1;
                else if (position > pagecount)
                    position = pagecount;

                PreviewImage.Source = _report.GetBitmapDecoder().Frames[position - 1];
                UpdateToolBarButton();
            }
            finally
            {
                Application.Current.MainWindow.Cursor = null;
            }
        }

        private void ExportMethod(ReportType rType)
        {
            using var saveFileDialog1 = new System.Windows.Forms.SaveFileDialog
            {
                RestoreDirectory = true,
            };

            switch (rType)
            {
                case ReportType.Pdf:
                    saveFileDialog1.Filter = @"Adobe PDF (*.pdf)|*.pdf";
                    saveFileDialog1.FileName = _report.Report.DisplayName + ".pdf";
                    break;
                case ReportType.Excel:
                    saveFileDialog1.Filter = @"Microsoft Excel (*.xls)|*.xls";
                    saveFileDialog1.FileName = _report.Report.DisplayName + ".xls";
                    break;
                case ReportType.Word:
                    saveFileDialog1.Filter = @"Microsoft Word (*.doc)|*.doc";
                    saveFileDialog1.FileName = _report.Report.DisplayName + ".doc";
                    break;
                case ReportType.Image:
                    saveFileDialog1.Filter = @"Image PNG (*.png)|*.png";
                    saveFileDialog1.FileName = _report.Report.DisplayName + ".png";
                    break;
            }

            if (saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                _report.Path = saveFileDialog1.FileName;
                _report.Reporttype = rType;
                _report.Print();
            }
        }

        /// <summary>
        /// Previous page
        /// </summary>
        private void PreviousImage_Click(object sender, RoutedEventArgs e)
        {
            if (_pos > 0)
                PageSpinner.Value = _pos - 1;
        }

        /// <summary>
        /// Next page
        /// </summary>
        private void NextImage_Click(object sender, RoutedEventArgs e)
        {
            if (_pos < _report.PagesCount)
                PageSpinner.Value = _pos + 1;
        }

        /// <summary>
        /// First page
        /// </summary>
        private void FirstImage_Click(object sender, RoutedEventArgs e)
        {
            PageSpinner.Value = 0;
        }

        /// <summary>
        /// Last page
        /// </summary>        
        private void LastImage_Click(object sender, RoutedEventArgs e)
        {
            PageSpinner.Value = _report.PagesCount;
        }

        /// <summary>
        /// Export to word file format button
        /// </summary>        
        private void MenuItemWord_Click(object sender, RoutedEventArgs e) => ExportMethod(ReportType.Word);

        /// <summary>
        /// Export to excel file format button
        /// </summary>        
        private void MenuItemExcel_Click(object sender, RoutedEventArgs e) => ExportMethod(ReportType.Excel);

        /// <summary>
        /// Export to PNG file format button
        /// </summary>        
        private void MenuItemPNG_Click(object sender, RoutedEventArgs e) => ExportMethod(ReportType.Image);
        
        /// <summary>
        /// Export to PDF file format button
        /// </summary>        
        private void ExportDefault_Click(object sender, RoutedEventArgs e) => ExportMethod(ReportType.Pdf);

        /// <summary>
        /// Call a RDLCPrintDialog
        /// </summary>
        private void TBBPrintWithProperties_Click(object sender, RoutedEventArgs e)
        {
            if (_report == null) return;

            var printerDialog = new RdlcPrinterDialog
            {
                Report = _report
            };

            printerDialog.ShowAsync();
        }

        private void PerCent50Button_Click(object sender, RoutedEventArgs e) => ZoomSlider.Value = 50;

        private void PerCent100Button_Click(object sender, RoutedEventArgs e) => ZoomSlider.Value = 100;

        private void PerCent150Button_Click(object sender, RoutedEventArgs e) => ZoomSlider.Value = 150;

        private void PerCent200Button_Click(object sender, RoutedEventArgs e) => ZoomSlider.Value = 200;

        private void PerCent250Button_Click(object sender, RoutedEventArgs e) => ZoomSlider.Value = 250;

        private void CommandBinding_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            if (_report == null) return;
        }

        private void OnSpinnerChanged(NumberBox sender, NumberBoxValueChangedEventArgs args)
        {
            if (_pos == (int)sender.Value || _report == null)
                return;

            _pos = (int)sender.Value;
            ChangeImage(_pos);
        }

        private void ZoomPopupButton_Click(SplitButton sender, SplitButtonClickEventArgs args)
        {
            ZoomSlider.Value = 100;
        }

        private void UpdateZoomText(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            ZoomValueTextBloc.Text = $"{e.NewValue:N0}";
        }

        private void ZoomByMouseWheel(object sender, MouseWheelEventArgs e)
        {
            if (Keyboard.Modifiers != ModifierKeys.Control)
                return;

            if (e.Delta > 0)
                ZoomSlider.Value += ZoomSlider.TickFrequency;

            else if (e.Delta < 0)
                ZoomSlider.Value -= ZoomSlider.TickFrequency;
        }

        private void ExportMenu_Click(SplitButton sender, SplitButtonClickEventArgs args)
        {
            ExportDefault_Click(sender, new RoutedEventArgs());
        }
    }
}
