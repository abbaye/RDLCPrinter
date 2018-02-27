using System;
using System.Windows;
using DSoft;
using Microsoft.Reporting.WinForms;

namespace RDLCDemo
{
    /// <summary>
    /// Logique d'interaction pour MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        NorthwindDataSetTableAdapters.ProductsByCategoriesTableAdapter _dataAdapter;
        NorthwindDataSet _northWindDataSet;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            LoadDataReport();
        }

        private void LoadDataReport()
        {
            //Create the local report
            var report = new LocalReport();
            report.ReportEmbeddedResource = "RDLCDemo.ReportTest.rdlc";


            //Create the dataset            
            _northWindDataSet = new NorthwindDataSet();
            _dataAdapter = new NorthwindDataSetTableAdapters.ProductsByCategoriesTableAdapter();

            _northWindDataSet.DataSetName = "NorthwindDataSet";
            _northWindDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            _dataAdapter.ClearBeforeFill = true;

            //Created datasource and binding source
            var dataSource = new ReportDataSource();
            var source = new System.Windows.Forms.BindingSource();

            dataSource.Name = "ProductsDataSources";  //the datasource name in the RDLC report
            dataSource.Value = source;
            source.DataMember = "ProductsByCategories";
            source.DataSource = _northWindDataSet;
            report.DataSources.Add(dataSource);


            //Fill Data in the dataset
            _dataAdapter.Fill(_northWindDataSet.ProductsByCategories);

            //Create the printer/export rdlc printer
            var rdlcPrinter = new RDLCPrinter(report);

            rdlcPrinter.BeforeRefresh += rdlcPrinter_BeforeRefresh;

            //Load in report viewer
            ReportViewer.Report = rdlcPrinter;
        }

        private void rdlcPrinter_BeforeRefresh(object sender, EventArgs e)
        {
            //Fill Data in the dataset            
            _dataAdapter.Fill(_northWindDataSet.ProductsByCategories);
        }
    }
}
