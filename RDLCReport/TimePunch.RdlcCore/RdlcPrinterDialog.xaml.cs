//////////////////////////////////////////////
// MIT - 2012-2020
// Author : Derek Tremblay (derektremblay666@gmail.com)
// Contributor : Martin Savard (2013)
// https://github.com/abbaye/RDLCPrinter
//////////////////////////////////////////////

using System.Drawing.Printing;
using System.Printing;
using System.Windows.Input;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Controls_ContentDialog = Microsoft.UI.Xaml.Controls.ContentDialog;
using Controls_ContentDialogButtonClickEventArgs = Microsoft.UI.Xaml.Controls.ContentDialogButtonClickEventArgs;

namespace TimePunch.Rdlc
{
    /// <summary>
    /// RDLC Printer settings Dialog
    /// </summary>
    public partial class RdlcPrinterDialog
    {
        private RdlcPrinter? _report;
        private PrintQueue _currentPrinter = LocalPrintServer.GetDefaultPrintQueue();

        private readonly PrintDocument _printer = new();
        private readonly LocalPrintServer _printServer = new();

        public RdlcPrinter? Report
        {
            get => _report;

            set
            {
                _report = value;

                if (_report is { IsDefaultLandscape: true })
                    _printer.DefaultPageSettings.Landscape = true;
            }
        }

        public string PrintText 
        {
            get
            {
                var resourceLoader = Windows.ApplicationModel.Resources.ResourceLoader.GetForCurrentView("Resources");
                return resourceLoader.GetString("Print");
            }
        }
        public string CloseText 
        {
            get
            {
                var resourceLoader = Windows.ApplicationModel.Resources.ResourceLoader.GetForCurrentView("Resources");
                return resourceLoader.GetString("Close");
            }
        }

        public RdlcPrinterDialog() => InitializeComponent();

        private async void Window_Loaded(object sender, RoutedEventArgs routedEventArgs) => await RefreshWindow();

        private async Task RefreshWindow()
        {
            if (_report == null) 
                return;

            if (_report.CopyNumber >= 1)
                NumberOfCopySpinner.Value = _report.CopyNumber;

            //update spinner
            var pageCount = await _report.GetPagesCountAsync();
            FirstPageSpinner.Maximum = pageCount;
            LastPageSpinner.Maximum = pageCount;

            FirstPageSpinner.Value = 1;
            LastPageSpinner.Value = pageCount;

            //Get all printer
            PrinterName.ItemsSource = _printServer.GetPrintQueues(new[] { EnumeratedPrintQueueTypes.Local, EnumeratedPrintQueueTypes.Connections });
            PrinterName.DisplayMemberPath = "FullName";
            PrinterName.SelectedValue = "FullName";

            //Select Default printer
            for (var i = 0; i < PrinterName.Items.Count; i++)
            {
                var testPrint = (PrintQueue)PrinterName.Items[i];
                if (testPrint.FullName == _currentPrinter.FullName)
                    PrinterName.SelectedIndex = i;
            }

            UpdatePrinterState();
        }

        /// <summary>
        /// Select user printer
        /// </summary>
        private void CboImprimanetNom_SelectionChanged(object sender, SelectionChangedEventArgs selectionChangedEventArgs)
        {
            PrinterState.Text = "";
            _currentPrinter = (PrintQueue)PrinterName.SelectedItem;
            UpdatePrinterState();
        }

        /// <summary>
        /// Launch print
        /// </summary>
        private async void OK_Click(ContentDialog sender, ContentDialogButtonClickEventArgs contentDialogButtonClickEventArgs)
        {
            _printer.PrinterSettings.PrinterName = _currentPrinter.FullName;
            if (AllPages.IsChecked == true)
                _printer.PrinterSettings.PrintRange = PrintRange.AllPages;
            else
            {
                _printer.PrinterSettings.PrintRange = PrintRange.Selection;
                _printer.PrinterSettings.FromPage = (int)FirstPageSpinner.Value;
                _printer.PrinterSettings.ToPage = (int)LastPageSpinner.Value;
            }

            if (Report != null)
            {
                Report.SetPrintDoc(_printer);
                Report.CopyNumber = Math.Max(0, (int)NumberOfCopySpinner.Value);
                await Report.Print();
            }
        }

        private void UpdatePrinterState()
        {
            PrinterState.Text = _currentPrinter.IsNotAvailable == false ? "Ready" : "Offline";
        }
    }
}
