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
using System.Windows.Navigation;
using System.Windows.Shapes;

using DSoft;
using Microsoft.Reporting.WinForms;

namespace RDLCDemo
{
    /// <summary>
    /// Logique d'interaction pour MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {            
            
            //Create the local report
            LocalReport report = new LocalReport();
            report.ReportEmbeddedResource = "RDLCDemo.ReportTest.rdlc";
            

            //Create the dataset            
            NorthwindDataSet northWindDataSet = new NorthwindDataSet();
            NorthwindDataSetTableAdapters.ProductsByCategoriesTableAdapter dataAdapter = new NorthwindDataSetTableAdapters.ProductsByCategoriesTableAdapter();

            northWindDataSet.DataSetName = "NorthwindDataSet";
            northWindDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            dataAdapter.ClearBeforeFill = true;

            //Created datasource and binding source
            ReportDataSource dataSource = new ReportDataSource();
            System.Windows.Forms.BindingSource source = new System.Windows.Forms.BindingSource();

            dataSource.Name = "ProductsDataSources";  //the datasource name in the RDLC report
            dataSource.Value = source;
            source.DataMember = "ProductsByCategories";
            source.DataSource = northWindDataSet;
            report.DataSources.Add(dataSource);


            //Fill Data in the dataset
            dataAdapter.Fill(northWindDataSet.ProductsByCategories);
            
            //Create the printer/export rdlc printer
            RDLCPrinter rdlcPrinter = new RDLCPrinter(report);

            //Load in report viewer
            ReportViewer.Report = rdlcPrinter;            
        }
    }
}
