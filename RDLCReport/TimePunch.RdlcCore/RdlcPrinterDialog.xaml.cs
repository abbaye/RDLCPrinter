//////////////////////////////////////////////
// MIT - 2012-2020
// Author : Derek Tremblay (derektremblay666@gmail.com)
// Contributor : Martin Savard (2013)
// https://github.com/abbaye/RDLCPrinter
//////////////////////////////////////////////

using System.Drawing.Printing;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using ResourceLoader = Microsoft.Windows.ApplicationModel.Resources.ResourceLoader;

namespace TimePunch.Rdlc
{
    /// <summary>
    /// RDLC Printer settings Dialog
    /// </summary>
    public partial class RdlcPrinterDialog
    {
        private RdlcPrinter? _report;
        private readonly PrintDocument _printer = new();
        private string _currentPrinter;
        private readonly ResourceLoader _resourceLoader = new(ResourceLoader.GetDefaultResourceFilePath(), "TpRdlcViewer/Resources");

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

        public string PrintText => _resourceLoader.GetString("Print/ToolTipService/ToolTip");

        public string CloseText => _resourceLoader.GetString("Close");

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
            _currentPrinter = new PrinterSettings().PrinterName;
            PrinterName.ItemsSource = PrinterSettings.InstalledPrinters; //_printServer.GetPrintQueues(new[] { EnumeratedPrintQueueTypes.Local, EnumeratedPrintQueueTypes.Connections });
            //PrinterName.DisplayMemberPath = "FullName";
            //PrinterName.SelectedValue = "FullName";

            //Select Default printer
            for (var i = 0; i < PrinterName.Items.Count; i++)
            {
                if (_currentPrinter.Equals(PrinterName.Items[i]))
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
            _currentPrinter = (string)PrinterName.SelectedItem;
            UpdatePrinterState();
        }

        /// <summary>
        /// Launch print
        /// </summary>
        private async void OK_Click(ContentDialog sender, ContentDialogButtonClickEventArgs contentDialogButtonClickEventArgs)
        {
            _printer.PrinterSettings.PrinterName = _currentPrinter;
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
            PrinterState.Text = "Ready"; //_currentPrinter.IsNotAvailable == false ? "Ready" : "Offline";
        }
    }
}
