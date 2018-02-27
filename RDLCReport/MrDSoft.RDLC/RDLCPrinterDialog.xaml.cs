using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;
using System.Printing;
using System.Drawing.Printing;

namespace DSoft.RDLC
{
    /// <summary>
    /// RDLCPrinterDialog
    /// <remarks>
    /// CREDIT : 2013-2018 Derek Tremblay (abbaye), 2013 Martin Savard
    /// https://github.com/abbaye/RDLCPrinter
    /// </remarks>
    /// </summary>
    public partial class RDLCPrinterDialog
    {
        private readonly PrintDocument _printer = new PrintDocument();
        private RDLCPrinter _report;
        private PrintQueue _currentPrinter;
        private readonly LocalPrintServer _printServer = new LocalPrintServer();
        //private List<PrintQueue> _printerList = new List<PrintQueue>();
        private string ImgSource; 

        public RDLCPrinter Report
        {
            get => _report;

            set
            {
                _report = value;

                if (_report.isDefaultLandscape == true)                
                    _printer.DefaultPageSettings.Landscape = true;   
            }
        }

        public RDLCPrinterDialog() => InitializeComponent();

        private void Window_Loaded(object sender, RoutedEventArgs e) => RefreshWindow();

        private void RefreshWindow()
        {
            if (_report == null) return;

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

            //Check if printer is ready
            if (_currentPrinter.IsNotAvailable == false)
            {
                lblImprimanteStatus.Content = "Ready";
                ImgSource = @"pack://application:,,,/RDLCPrinter;component/Resources/Button-Blank-Green.ico";

            }
            else
            {
                lblImprimanteStatus.Content = "Offline";
                ImgSource = @"pack://application:,,,/RDLCPrinter;component/Resources/Button-Blank-Red.ico";

            }
            ReadyImage.Source = new BitmapImage(new Uri(ImgSource));
        }

        /// <summary>
        /// Select user printer
        /// </summary>
        private void cboImprimanetNom_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            lblImprimanteStatus.Content = "";

            _currentPrinter = (PrintQueue)cboImprimanteNom.SelectedItem;

            if (_currentPrinter.IsNotAvailable == false)
            {
                lblImprimanteStatus.Content = "Ready";                
                _printer.PrinterSettings.PrinterName = _currentPrinter.FullName;
                lblEmplacementImprimante.Content = _currentPrinter.QueuePort.Name;
                ImgSource = @"pack://application:,,,/RDLCPrinter;component/Resources/Button-Blank-Green.ico";
            }
            else
            {
                lblImprimanteStatus.Content = "Offline";
                ImgSource = @"pack://application:,,,/RDLCPrinter;component/Resources/Button-Blank-Red.ico";
            }

            ReadyImage.Source = new BitmapImage(new Uri(ImgSource));
        }

        /// <summary>
        /// Launch print
        /// </summary>
        private void OK_Click(object sender, RoutedEventArgs e)
        {            
            PreparePrint();
            Report.PrintDoc = _printer;

            Report.CopyNumber = NumberOfCopySpinner.Value.HasValue ? NumberOfCopySpinner.Value.Value : 1;

            Report.Print();
            Close();
        }

        /// <summary>
        /// Page to page for printing
        /// </summary>
        private void PreparePrint()
        {
            if (cmdAllPageButton.IsChecked != false) return;

            _printer.PrinterSettings.FromPage = FirstPageSpinner.Value.Value;
            _printer.PrinterSettings.ToPage = LastPageSpinner.Value.Value;
        }


        /// <summary>
        /// Close window 
        /// </summary>        
        private void Annuler_Click(object sender, RoutedEventArgs e) => Close();

        /////// <summary>
        /////// Cancel ctrl-v (paste)
        /////// </summary>        
        ////private void OnCancelCommand(object sender, DataObjectEventArgs e)
        ////{
        ////    e.CancelCommand();
        ////}

        private void cmdAllPageButton_Click(object sender, RoutedEventArgs e)
        {
            if (cmdAllPageButton.IsChecked == true)
            {
                PageChoiceStackPanel.IsEnabled = false;
                FirstPageSpinner.Value = 1;
                LastPageSpinner.Value = _report.PagesCount;
            }
            else
                PageChoiceStackPanel.IsEnabled = true;
        }
    }
}
