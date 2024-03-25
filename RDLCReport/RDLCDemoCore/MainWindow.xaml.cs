using Microsoft.Reporting.NETCore;
using Microsoft.UI.Xaml;
using RDLCDemo.NorthwindDataSetTableAdapters;
using TimePunch.Rdlc;

namespace RDLCDemo
{
    /// <summary>
    /// Logique d'interaction pour MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow() => InitializeComponent();

        private void Window_Loaded(object sender, WindowActivatedEventArgs args) => LoadDataReport();

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

            var source = northWindDataSet.ProductsByCategories;

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
            RDLCReportViewer.Report = new RdlcPrinter(report);
        }
    }
}
