//////////////////////////////////////////////
// MIT - 2012-2020
// Author : Derek Tremblay (derektremblay666@gmail.com)
// Contributor : Martin Savard (2013)
// https://github.com/abbaye/RDLCPrinter
//////////////////////////////////////////////

using System.Diagnostics;
using Windows.Graphics.Imaging;
using Windows.System;
using Windows.UI.Core;
using Microsoft.UI.Input;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Controls.Primitives;
using PointerRoutedEventArgs = Microsoft.UI.Xaml.Input.PointerRoutedEventArgs;
using SplitButton = Microsoft.UI.Xaml.Controls.SplitButton;
using SplitButtonClickEventArgs = Microsoft.UI.Xaml.Controls.SplitButtonClickEventArgs;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media.Imaging;
using TimePunch.Rdlc.MethodExtension;
using WinRT.Interop;

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
       
        private Mutex? _onlyRefreshOnce;
        private readonly string _rdlcReportMutexName = Guid.NewGuid().ToString();
        
        public RdlcReportViewer()
        {
            InitializeComponent();
#pragma warning disable CS4014 // Because this call is not awaited, execution of the current method continues before the call is completed
            UpdateToolBarButton();
#pragma warning restore CS4014 // Because this call is not awaited, execution of the current method continues before the call is completed
        }

        /// <summary>
        /// Call refresh button
        /// </summary>        
        private async Task CmdRefresh_Click(object sender, RoutedEventArgs e) => await RefreshControl();

        public bool StartAfterExport
        {
            get => (bool)GetValue(StartAfterExportProperty);
            set => SetValue(StartAfterExportProperty, value);
        }

        /// <summary>
        /// Get or Set the RDLC report
        /// </summary>
        public RdlcPrinter? Report
        {
            get => (RdlcPrinter?)GetValue(ReportProperty);
            set => SetValue(ReportProperty, value);
        }
        private static async void OnReportChanged(DependencyObject dependencyObject, DependencyPropertyChangedEventArgs dependencyPropertyChangedEventArgs)
        {
            if (dependencyObject is RdlcReportViewer viewer)
            {
                await viewer.RefreshControl();
                await viewer.GiveFocus();
            }
        }

        public int Page
        {
            get => (int)GetValue(PageProperty);
            set => SetValue(PageProperty, value);
        }

        private static async void OnPageChanged(DependencyObject dependencyObject, DependencyPropertyChangedEventArgs e)
        {
            if (dependencyObject is RdlcReportViewer viewer)
                await viewer.UpdateImage();
        }

        /// <summary>
        /// Give the focus to usercontrol
        /// </summary>
        public async Task GiveFocus()
        {
            await FocusManager.TryFocusAsync(PreviewImage, FocusState.Programmatic);
        }

        /// <summary>
        /// Refresh logic.         
        /// </summary>
        public async Task RefreshControl()
        {
            if (Report == null) return;

            _onlyRefreshOnce = MutexAcl.Create(true, $"RDLCReport: {_rdlcReportMutexName}", out var lockWasTaken, null);
            if (!lockWasTaken)
                return;

            try
            {
                Report.Refresh();

                await LoadImage();

                Page = 1;
                PageSpinner.Maximum = await Report.GetPagesCountAsync();
                PageSpinner.Value = Page;
            }
            finally
            {
                _onlyRefreshOnce.ReleaseMutex();
            }
        }

        /// <summary>
        /// Update the toolbar button
        /// </summary>
        private async Task UpdateToolBarButton()
        {
            if (Page < 0) 
                return;

            if (Report != null)
            {
                ButtonExtention.EnableButton(TbbOpenInExcel);
                ExportMenu.IsEnabled = true;
                ExportMenu.Opacity = 1;
                ButtonExtention.EnableButton(TbbPrintWithProperties);
                ZoomInfoStackPanel.Visibility = Visibility.Visible;
                ZoomPopupButton.Visibility = Visibility.Visible;
                ZoomPopupButton.IsEnabled = true;
                ZoomPopupButton.Opacity = 1;
            }
            else
            {
                ButtonExtention.DisableButton(TbbOpenInExcel);
                ExportMenu.IsEnabled = false;
                ExportMenu.Opacity = 0.5;
                ButtonExtention.DisableButton(TbbPrintWithProperties);
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
                if (Report != null && Page == await Report.GetPagesCountAsync())
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

            if (Report != null && await Report.GetPagesCountAsync() > 1 && Page>0)
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
        private async Task LoadImage()
        {
            if (Page == 0)
            {
                await UpdateToolBarButton();
                PreviousImage.IsEnabled = false;
                PreviousImage.Opacity = 0.5;
            }
        }

        /// <summary>
        /// Chage page to the position
        /// </summary>
        private async Task UpdateImage()
        {
            if (Report == null)
                return;
            
            try
            {
                ProtectedCursor = InputCursor.CreateFromCoreCursor(new CoreCursor(CoreCursorType.Wait, 0));

                int position = Page;
                var pageCount = await Report.GetPagesCountAsync();

                //Check interval
                if (position <= 0)
                    position = 1;
                else if (position > pageCount)
                    position = pageCount;

                var decoder = await Report.GetBitmapDecoderAsync();
                var frame = await decoder?.GetFrameAsync((uint)position - 1);
                var bitmap = await frame.GetSoftwareBitmapAsync(frame.BitmapPixelFormat, BitmapAlphaMode.Premultiplied);
                var source = new SoftwareBitmapSource();
                await source.SetBitmapAsync(bitmap);

                PreviewImage.Source = source;
                await UpdateToolBarButton();
            }
            finally
            {
                ProtectedCursor = null;
            }
        }

        private async Task ExportMethod(ReportType rType)
        {
            if (Report == null)
                return;

            var saveFileDialog1 = new Windows.Storage.Pickers.FileSavePicker()
            {
                SuggestedStartLocation = Windows.Storage.Pickers.PickerLocationId.DocumentsLibrary,
            };
            InitializeWithWindow.Initialize(saveFileDialog1, Process.GetCurrentProcess().MainWindowHandle);

            switch (rType)
            {
                case ReportType.Pdf:
                    saveFileDialog1.FileTypeChoices.Add("Adobe PDF (*.pdf)", [ ".pdf" ]);
                    saveFileDialog1.SuggestedFileName = Report.Report.DisplayName + ".pdf";
                    break;
                case ReportType.Excel:
                    saveFileDialog1.FileTypeChoices.Add("Microsoft Excel (*.xlsx)", [".xlsx" ]);
                    saveFileDialog1.SuggestedFileName = Report.Report.DisplayName + ".xlsx";
                    break;
                case ReportType.Word:
                    saveFileDialog1.FileTypeChoices.Add("Microsoft Word (*.docx)", [".docx"]);
                    saveFileDialog1.SuggestedFileName = Report.Report.DisplayName + ".docx";
                    break;
                case ReportType.Image:
                    saveFileDialog1.FileTypeChoices.Add("Image PNG (*.png)", [".png"]);
                    saveFileDialog1.SuggestedFileName = Report.Report.DisplayName + ".png";
                    break;
            }

            var result = await saveFileDialog1.PickSaveFileAsync();

            if (result != null)
            {
                Report.Path = result.Path;
                Report.ReportType = rType;
                await Report.Print();

                if (StartAfterExport)
                {
                    var processStart = new ProcessStartInfo(result.Path){ UseShellExecute = true };
                    Process.Start(processStart);
                }
            }
        }

        /// <summary>
        /// Previous page
        /// </summary>
        private void PreviousImage_Click(object sender, RoutedEventArgs routedEventArgs)
        {
            if (Page > 0)
                PageSpinner.Value = Page - 1;
        }

        /// <summary>
        /// Next page
        /// </summary>
        private async void NextImage_Click(object sender, RoutedEventArgs routedEventArgs)
        {
            if (Report == null)
                return;

            if (Page < await Report.GetPagesCountAsync())
                PageSpinner.Value = Page + 1;
        }

        /// <summary>
        /// First page
        /// </summary>
        private void FirstImage_Click(object sender, RoutedEventArgs routedEventArgs)
        {
            PageSpinner.Value = 0;
        }

        /// <summary>
        /// Last page
        /// </summary>        
        private async void LastImage_Click(object sender, RoutedEventArgs routedEventArgs)
        {
            if (Report == null)
                return;

            PageSpinner.Value = await Report.GetPagesCountAsync();
        }

        /// <summary>
        /// Export to word file format button
        /// </summary>        
        private async void MenuItemWord_Click(object sender, RoutedEventArgs routedEventArgs) => await ExportMethod(ReportType.Word);

        /// <summary>
        /// Export to excel file format button
        /// </summary>        
        private async void MenuItemExcel_Click(object sender, RoutedEventArgs routedEventArgs) => await ExportMethod(ReportType.Excel);

        /// <summary>
        /// Export to PNG file format button
        /// </summary>        
        private async void MenuItemPNG_Click(object sender, RoutedEventArgs routedEventArgs) => await ExportMethod(ReportType.Image);
        
        /// <summary>
        /// Export to PDF file format button
        /// </summary>        
        private async void ExportDefault_Click(object sender, RoutedEventArgs routedEventArgs) => await ExportMethod(ReportType.Pdf);

        /// <summary>
        /// Call a RDLCPrintDialog
        /// </summary>
        private async void TBBPrintWithProperties_Click(object sender, RoutedEventArgs routedEventArgs)
        {
            if (Report == null) return;

            var printerDialog = new RdlcPrinterDialog
            {
                Report = Report,
                XamlRoot = XamlRoot
            };

            await printerDialog.ShowAsync();
        }

        private async void CmdExcel_Click(object sender, RoutedEventArgs routedEventArgs)
        {
            if (Report == null) return;

            var fileName = Path.Combine(Path.GetTempPath(), Path.ChangeExtension(Path.GetRandomFileName(), ".xlsx"));
            Report.Path = fileName;
            Report.ReportType = ReportType.Excel;
            await Report.Print();

            var processStart = new ProcessStartInfo(fileName){ UseShellExecute = true };
            Process.Start(processStart);
        }

        private void PerCent50Button_Click(object sender, RoutedEventArgs routedEventArgs) => ZoomSlider.Value = 50;

        private void PerCent100Button_Click(object sender, RoutedEventArgs routedEventArgs) => ZoomSlider.Value = 100;

        private void PerCent150Button_Click(object sender, RoutedEventArgs routedEventArgs) => ZoomSlider.Value = 150;

        private void PerCent200Button_Click(object sender, RoutedEventArgs routedEventArgs) => ZoomSlider.Value = 200;

        private void PerCent250Button_Click(object sender, RoutedEventArgs routedEventArgs) => ZoomSlider.Value = 250;

        private void CommandBinding_Executed(object sender, RoutedEventArgs e)    
        {
            if (Report == null) return;
        }

        private void OnSpinnerChanged(NumberBox sender, NumberBoxValueChangedEventArgs args)
        {
            if (Page == (int)sender.Value || Report == null)
                return;

            Page = (int)sender.Value;
        }

        private void ZoomPopupButton_Click(SplitButton splitButton, SplitButtonClickEventArgs splitButtonClickEventArgs)
        {
            ZoomSlider.Value = 100;
        }

        private void UpdateZoomText(object sender, RangeBaseValueChangedEventArgs rangeBaseValueChangedEventArgs)
        {
            if (ScrollViewerControl is null)
                return;

            ZoomValueTextBloc.Text = $"{rangeBaseValueChangedEventArgs.NewValue:N0}";

            double zoomFactor = rangeBaseValueChangedEventArgs.NewValue / 100.0;
            ScrollViewerControl.ChangeView(
                horizontalOffset: null,
                verticalOffset: null,
                zoomFactor: (float?)zoomFactor, false);
        }

        private void ZoomByMouseWheel(object sender, PointerRoutedEventArgs pointerRoutedEventArgs)
        {
            if (pointerRoutedEventArgs.KeyModifiers != VirtualKeyModifiers.Control)
                return;

            var delta = pointerRoutedEventArgs.GetCurrentPoint((UIElement)sender).Properties.MouseWheelDelta;

            switch (delta)
            {
                case > 0:
                    ZoomSlider.Value += delta/10.0;
                    break;
                
                case < 0:
                    ZoomSlider.Value += delta/10.0;
                    break;
            }
        }

        private void ExportMenu_Click(Microsoft.UI.Xaml.Controls.SplitButton sender, SplitButtonClickEventArgs splitButtonClickEventArgs)
        {
            ExportDefault_Click(sender, new RoutedEventArgs());
        }
    }
}
