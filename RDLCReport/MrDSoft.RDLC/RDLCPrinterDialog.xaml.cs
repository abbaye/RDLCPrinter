//////////////////////////////////////////////
// MIT - 2012-2020
// Author : Derek Tremblay (derektremblay666@gmail.com)
// Contributor : Martin Savard (2013)
// https://github.com/abbaye/RDLCPrinter
//////////////////////////////////////////////

using System;
using System.Drawing.Printing;
using System.Printing;
using System.Windows;
using System.Windows.Controls;
using ModernWpf.Controls;

namespace TimePunch.Rdlc
{
    /// <summary>
    /// RDLC Printer settings Dialog
    /// </summary>
    public partial class RdlcPrinterDialog
    {
        private RdlcPrinter _report;
        private PrintQueue _currentPrinter;

        private readonly PrintDocument _printer = new();
        private readonly LocalPrintServer _printServer = new();

        public RdlcPrinter Report
        {
            get => _report;

            set
            {
                _report = value;

                if (_report.IsDefaultLandscape == true)
                    _printer.DefaultPageSettings.Landscape = true;
            }
        }

        public RdlcPrinterDialog() => InitializeComponent();

        private void Window_Loaded(object sender, RoutedEventArgs e) => RefreshWindow();

        private void RefreshWindow()
        {
            if (_report == null) 
                return;

            if (_report.CopyNumber >= 1)
                NumberOfCopySpinner.Value = _report.CopyNumber;

            //update spinner
            FirstPageSpinner.Maximum = _report.PagesCount;
            LastPageSpinner.Maximum = _report.PagesCount;

            FirstPageSpinner.Value = 1;
            LastPageSpinner.Value = _report.PagesCount;

            //Get all printer
            cboImprimanteNom.ItemsSource = _printServer.GetPrintQueues(new[] { EnumeratedPrintQueueTypes.Local, EnumeratedPrintQueueTypes.Connections });
            cboImprimanteNom.DisplayMemberPath = "FullName";
            cboImprimanteNom.SelectedValue = "FullName";

            //Select Default printer
            _currentPrinter = LocalPrintServer.GetDefaultPrintQueue();
            for (var i = 0; i < cboImprimanteNom.Items.Count; i++)
            {
                var testPrint = (PrintQueue)cboImprimanteNom.Items[i];
                if (testPrint.FullName == _currentPrinter.FullName)
                    cboImprimanteNom.SelectedIndex = i;
            }

            UpdatePrinterState();
        }

        /// <summary>
        /// Select user printer
        /// </summary>
        private void CboImprimanetNom_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ImprimanteStatus.Content = "";
            _currentPrinter = (PrintQueue)cboImprimanteNom.SelectedItem;
            UpdatePrinterState();
        }

        /// <summary>
        /// Launch print
        /// </summary>
        private void OK_Click(ContentDialog contentDialog, ContentDialogButtonClickEventArgs args)
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

            Report.PrintDoc = _printer;
            Report.CopyNumber = Math.Max(0, (int)NumberOfCopySpinner.Value);
            Report.Print();
        }

        private void UpdatePrinterState()
        {
            if (_currentPrinter.IsNotAvailable == false)
                ImprimanteStatus.Content = "Ready";
            else
                ImprimanteStatus.Content = "Offline";
        }

    }
}
