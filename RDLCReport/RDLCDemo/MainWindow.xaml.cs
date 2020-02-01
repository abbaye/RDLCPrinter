using System.Windows;
using System.Windows.Forms;
using DSoft;
using Microsoft.Reporting.WinForms;
using RDLCDemo.NorthwindDataSetTableAdapters;

namespace RDLCDemo
{
    /// <summary>
    /// Logique d'interaction pour MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow() => InitializeComponent();

        private void Window_Loaded(object sender, RoutedEventArgs e) => LoadDataReport();

        private void LoadDataReport()
        {
            //Create the local report
            var report = new LocalReport
            {
                ReportEmbeddedResource = "RDLCDemo.ReportTest.rdlc"
            };

            //Create object
            var northWindDataSet = new NorthwindDataSet
            {
                DataSetName = nameof(NorthwindDataSet),
                SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
            };

            var dataAdapter = new ProductsByCategoriesTableAdapter
            {
                ClearBeforeFill = true
            };

            var source = new BindingSource
            {
                DataMember = nameof(northWindDataSet.ProductsByCategories),
                DataSource = northWindDataSet
            };

            //Created datasource and binding source
            var dataSource = new ReportDataSource
            {
                Name = "ProductsDataSources",  //the datasource name in the RDLC report
                Value = source
            };

            //Add data source
            report.DataSources.Add(dataSource);
            
            //Fill Data in the dataset
            dataAdapter.Fill(northWindDataSet.ProductsByCategories);

            //Load in report viewer
            RDLCReportViewer.Report = new RDLCPrinter(report);
        }
    }
}
