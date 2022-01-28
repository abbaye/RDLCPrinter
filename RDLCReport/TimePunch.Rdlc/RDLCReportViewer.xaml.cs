//////////////////////////////////////////////
// MIT - 2012-2020
// Author : Derek Tremblay (derektremblay666@gmail.com)
// Contributor : Martin Savard (2013)
// https://github.com/abbaye/RDLCPrinter
//////////////////////////////////////////////

using System.Diagnostics;
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
        public static readonly DependencyProperty ReportProperty = DependencyProperty.Register("Report", typeof(RdlcPrinter), typeof(RdlcReportViewer), new PropertyMetadata(null, OnReportChanged));
        public static readonly DependencyProperty PageProperty = DependencyProperty.Register("Page", typeof(int), typeof(RdlcReportViewer), new PropertyMetadata(0, OnPageChanged));
        public static readonly DependencyProperty StartAfterExportProperty = DependencyProperty.Register("StartAfterExport", typeof(bool), typeof(RdlcReportViewer), new PropertyMetadata(true));

        public RdlcReportViewer()
        {
            InitializeComponent();
            UpdateToolBarButton();
        }

        /// <summary>
        /// Call refresh button
        /// </summary>        
        private void CmdRefresh_Click(object sender, RoutedEventArgs e) => RefreshControl();

        public bool StartAfterExport
        {
            get => (bool)GetValue(StartAfterExportProperty);
            set => SetValue(StartAfterExportProperty, value);
        }

        /// <summary>
        /// Get or Set the RDLC report
        /// </summary>
        public RdlcPrinter Report
        {
            get => (RdlcPrinter)GetValue(ReportProperty);
            set => SetValue(ReportProperty, value);
        }
        private static void OnReportChanged(DependencyObject dependencyObject, DependencyPropertyChangedEventArgs dependencyPropertyChangedEventArgs)
        {
            if (dependencyObject is RdlcReportViewer viewer)
            {
                viewer.RefreshControl();
                viewer.GiveFocus();
            }
        }

        public int Page
        {
            get => (int)GetValue(PageProperty);
            set => SetValue(PageProperty, value);
        }

        private static void OnPageChanged(DependencyObject dependencyObject, DependencyPropertyChangedEventArgs e)
        {
            if (dependencyObject is RdlcReportViewer viewer)
                viewer.UpdateImage();
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
            if (Report == null) return;

            Report.Refresh();

            DispatcherHelper.DoEvents(); //Clear the ui message queud

            LoadImage();

            Page = 1;
            PageSpinner.Maximum = Report.PagesCount;
            PageSpinner.Value = Page;
            UpdateImage();
        }

        /// <summary>
        /// Update the toolbar button
        /// </summary>
        private void UpdateToolBarButton()
        {
            if (Page < 0) 
                return;

            if (Report != null)
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

            if (Page == 1)
            {
                PagerSeparator.Visibility = Visibility.Visible;
                ButtonExtention.DisableButton(PreviousImage);
                ButtonExtention.DisableButton(FirstImage);
                ButtonExtention.EnableButton(NextImage);
                ButtonExtention.EnableButton(LastImage);
            }
            else
            {
                if (Report != null && Page == Report.PagesCount)
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

            if (Report != null && Report.PagesCount > 1 && Page>0)
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
            if (Page == 0)
            {
                UpdateToolBarButton();
                PreviousImage.IsEnabled = false;
                PreviousImage.Opacity = 0.5;
            }
        }

        /// <summary>
        /// Chage page to the position
        /// </summary>
        private void UpdateImage()
        {
            if (Report == null || Application.Current.MainWindow == null)
                return;

            try
            {
                Application.Current.MainWindow.Cursor = Cursors.Wait;

                int position = Page;
                var pagecount = Report.GetBitmapDecoder().Frames.Count;

                //Check interval
                if (position <= 0)
                    position = 1;
                else if (position > pagecount)
                    position = pagecount;

                PreviewImage.Source = Report.GetBitmapDecoder().Frames[position - 1];
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
                    saveFileDialog1.FileName = Report.Report.DisplayName + ".pdf";
                    break;
                case ReportType.Excel:
                    saveFileDialog1.Filter = @"Microsoft Excel (*.xls)|*.xls";
                    saveFileDialog1.FileName = Report.Report.DisplayName + ".xls";
                    break;
                case ReportType.Word:
                    saveFileDialog1.Filter = @"Microsoft Word (*.doc)|*.doc";
                    saveFileDialog1.FileName = Report.Report.DisplayName + ".doc";
                    break;
                case ReportType.Image:
                    saveFileDialog1.Filter = @"Image PNG (*.png)|*.png";
                    saveFileDialog1.FileName = Report.Report.DisplayName + ".png";
                    break;
            }

            if (saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                Report.Path = saveFileDialog1.FileName;
                Report.Reporttype = rType;
                Report.Print();

                if (StartAfterExport)
                {
                    var processStart = new ProcessStartInfo(saveFileDialog1.FileName);
                    Process.Start(processStart);
                }
            }
        }

        /// <summary>
        /// Previous page
        /// </summary>
        private void PreviousImage_Click(object sender, RoutedEventArgs e)
        {
            if (Page > 0)
                PageSpinner.Value = Page - 1;
        }

        /// <summary>
        /// Next page
        /// </summary>
        private void NextImage_Click(object sender, RoutedEventArgs e)
        {
            if (Page < Report.PagesCount)
                PageSpinner.Value = Page + 1;
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
            PageSpinner.Value = Report.PagesCount;
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
            if (Report == null) return;

            var printerDialog = new RdlcPrinterDialog
            {
                Report = Report
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
            if (Report == null) return;
        }

        private void OnSpinnerChanged(NumberBox sender, NumberBoxValueChangedEventArgs args)
        {
            if (Page == (int)sender.Value || Report == null)
                return;

            Page = (int)sender.Value;
            UpdateImage();
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
