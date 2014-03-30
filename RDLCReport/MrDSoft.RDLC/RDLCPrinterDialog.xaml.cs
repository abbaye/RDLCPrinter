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
using System.Windows.Shapes;
using System.Drawing;
using System.Printing;
using System.Drawing.Printing;
using Microsoft.Reporting.WinForms;
using System.IO;
using DSoft;
using DSoft.MethodExtension;
using DSoft.RDLCReport;

namespace DSoft.RDLC
{
    /// <summary>
    /// Interaction logic for RDLCPrinterDialog.xaml
    /// </summary>
    public partial class RDLCPrinterDialog : Window
    {
        private PrintDocument _printer = new PrintDocument();
        private RDLCPrinter _report = null;
        private int _nbrPageRapport = 1;
        private PrintQueue _currentPrinter = null;
        private LocalPrintServer _printServer = new LocalPrintServer();
        private List<PrintQueue> _printerList = new List<PrintQueue>();
        private string ImgSource; 

        public int NbrPageRapport
        {
            get
            {
                return this._nbrPageRapport;
            }
            set 
            {
                _nbrPageRapport = value;
            }
            
        }

        public RDLCPrinter CurrentReport
        {
            get
            {
                return _report;
            }

            set
            {
                _report = value;

                if (_report.isDefaultLandscape == true)
                {
                    _printer.DefaultPageSettings.Landscape = true;                    
                }
                else
                {
                    
                }
            }
        }

        public RDLCPrinterDialog()
        {
            InitializeComponent();
        }


        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            RefreshWindow();
            //disable Paste (coller) pour les TextBox du choix des pages
            DataObject.AddPastingHandler(FirstPage, this.OnCancelCommand);
            DataObject.AddPastingHandler(LastPage, this.OnCancelCommand);
        }

        private void RefreshWindow()
        {
            
            
            if (this._report.CopyNumber >= 1)
            {
                NumberOfCopySpinner.Value = this._report.CopyNumber;                
            }
            FirstPage.Text = "1";

            //obtien le nombre de page du rapport
            LastPage.Text = _nbrPageRapport.ToString();

            //Get all printer
            cboImprimanteNom.ItemsSource = _printServer.GetPrintQueues(new[] { EnumeratedPrintQueueTypes.Local, EnumeratedPrintQueueTypes.Connections }).Cast<PrintQueue>();
            cboImprimanteNom.DisplayMemberPath = "FullName";
            cboImprimanteNom.SelectedValue = "FullName";

            //Select Default printer
            _currentPrinter = LocalPrintServer.GetDefaultPrintQueue();
            for (int i = 0; i < cboImprimanteNom.Items.Count; i++)
            {
                PrintQueue testPrint = (PrintQueue)cboImprimanteNom.Items[i];
                if (testPrint.FullName.ToString() == _currentPrinter.FullName.ToString())
                {
                    cboImprimanteNom.SelectedIndex = i;
                }
            }

            //Check if printer is ready
            if (_currentPrinter.IsNotAvailable == false)
            {
                lblImprimanteStatus.Content = "Ready";
                ImgSource = @"pack://application:,,,/DSoft.RDLC;component/Resources/Button-Blank-Green.ico";

            }
            else
            {
                lblImprimanteStatus.Content = "Offline";
                ImgSource = @"pack://application:,,,/DSoft.RDLC;component/Resources/Button-Blank-Red.ico";

            }
            ReadyImage.Source = new BitmapImage(new Uri(ImgSource));
        }

        /// <summary>
        /// Select user printer
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cboImprimanetNom_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            

            lblImprimanteStatus.Content = "";

            _currentPrinter = (PrintQueue)cboImprimanteNom.SelectedItem;

            if (_currentPrinter.IsNotAvailable == false)
            {
                lblImprimanteStatus.Content = "Ready";                
                _printer.PrinterSettings.PrinterName = _currentPrinter.FullName;
                lblEmplacementImprimante.Content = _currentPrinter.QueuePort.Name;
                ImgSource = @"pack://application:,,,/DSoft.RDLC;component/Resources/Button-Blank-Green.ico";
            }
            else
            {
                lblImprimanteStatus.Content = "Offline";
                ImgSource = @"pack://application:,,,/DSoft.RDLC;component/Resources/Button-Blank-Red.ico";
            }

            ReadyImage.Source = new BitmapImage(new Uri(ImgSource));
        }

        /// <summary>
        /// Launch print
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OK_Click(object sender, RoutedEventArgs e)
        {            
            PreparePrint();
            CurrentReport.PrintDoc = this._printer;

            if (NumberOfCopySpinner.Value.HasValue)
                CurrentReport.CopyNumber = NumberOfCopySpinner.Value.Value;
            else
                CurrentReport.CopyNumber = 1;

            CurrentReport.Print();
            this.Close();
        }

        /// <summary>
        /// Page to page for printing
        /// </summary>
        /// <returns></returns>
        private void PreparePrint()
        {
            if (cmdAllPageButton.IsChecked == false)
            {
                _printer.PrinterSettings.FromPage = Convert.ToInt32(FirstPage.Text);
                _printer.PrinterSettings.ToPage = Convert.ToInt32(LastPage.Text);
            }
        }


        /// <summary>
        /// Close window 
        /// </summary>        
        private void Annuler_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void FirstPage_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.IsNumericKey())
                e.Handled = false;
            else
                e.Handled = true;
        }

        private void LastPage_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.IsNumericKey())
                e.Handled = false;
            else
                e.Handled = true;
        }

        /// <summary>
        /// Cancel ctrl-v (paste)
        /// </summary>        
        private void OnCancelCommand(object sender, DataObjectEventArgs e)
        {
            e.CancelCommand();
        }

        private void cmdAllPageButton_Click(object sender, RoutedEventArgs e)
        {
            if (cmdAllPageButton.IsChecked == true)
            {
                PageChoiceStackPanel.IsEnabled = false;
                FirstPage.Text = "1";
                LastPage.Text = _nbrPageRapport.ToString();
            }
            else
                PageChoiceStackPanel.IsEnabled = true;
        }

        private void FirstPage_LostFocus(object sender, RoutedEventArgs e)
        {
            if(FirstPage.Text.Trim() == "")
                FirstPage.Text = "1";

            if (Convert.ToInt32(FirstPage.Text) < 1 || Convert.ToInt32(FirstPage.Text) > _nbrPageRapport)
                FirstPage.Text = "1";
        }

        private void LastPage_LostFocus(object sender, RoutedEventArgs e)
        {
            if(LastPage.Text.Trim() == "")
                LastPage.Text = _nbrPageRapport.ToString();

            if (Convert.ToInt32(LastPage.Text) < 1 || Convert.ToInt32(LastPage.Text) > _nbrPageRapport)
                LastPage.Text = _nbrPageRapport.ToString();
        }

    }
}
