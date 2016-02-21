using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Windows.Media.Imaging;
using Microsoft.Reporting.WinForms;

namespace DSoft
{
    /// <summary>
    /// This class allow to Print and export RDLC Report. 
    /// <remarks>
    /// CREDIT : 2013-2016 Derek Tremblay (abbaye), 2013 Martin Savard
    /// https://github.com/abbaye/RDLCPrinter
    /// </remarks>
    /// </summary>
    public class RDLCPrinter : IDisposable
    {
        private int _currentPageIndex;
        private IList<Stream> _streams;
        private LocalReport _report;
        private int _Copies;
        private Metafile _pageImage = null;
        private Rectangle _adjustedRect = new Rectangle();
        private ReportType _ReportType;
        private string _path;
        private Warning[] _warnings = null;
        private string[] _streamids = null;
        private string _mimeType = string.Empty;
        private string _encoding = string.Empty;
        private string _extension = string.Empty;
        private string _filename = "";
        private PrintDocument _printDoc = null;
        private BitmapDecoder _dec = null;

        //Event
        public event EventHandler FileSaving;
        public event EventHandler FileSaved;
        public event EventHandler BeforeRefresh;
        
        /// <summary>
        /// Initialize an empty report object
        /// </summary>
        public RDLCPrinter()
        {

        }
        
        /// <summary>
        /// Initialize report object with default setting
        /// </summary>
        public RDLCPrinter(LocalReport report, ReportType rtype, int nbrPage, string path)
        {
            _report = report;
            _Copies = nbrPage;
            _ReportType = rtype;
            _path = path;
        }

        /// <summary>
        /// Initialize repot object with default setting
        /// Copies = 1
        /// ReportType = Printer
        /// Path = 0
        /// </summary>
        public RDLCPrinter(LocalReport report)
        {
            _report = report;
            _Copies = 1;
            _ReportType = ReportType.Printer;
            _path = "";
        }

        /// <summary>
        /// Refresh Report and Data
        /// </summary>
        public void Refresh()
        {
            if (_report != null)
            {
                if (BeforeRefresh != null)
                    BeforeRefresh(this, new EventArgs());

                _report.Refresh();
                _dec = null;
            }
        }
        
        #region Properties
        
        /// <summary>
        ///Get or set the printer for print report (if null, get default printer)
        /// </summary>        
        public PrintDocument PrintDoc
        {            
            get
            {
                return (_printDoc != null) ? _printDoc : GetDefaultPrinter();
            }
            set
            {
                _printDoc = value;
            }
        }


        /// <summary>
        /// get or set the type of report
        /// </summary>
        [DefaultValue(ReportType.Printer)]
        public ReportType Reporttype
        {
            get
            {
                return _ReportType;
            }
            set
            {
                _ReportType = value;
            }
        }

        /// <summary>
        /// Get the default orientation of report
        /// </summary>
        public bool? isDefaultLandscape
        {
            get
            {
                if (_report != null)
                    return _report.GetDefaultPageSettings().IsLandscape;
                else
                    return null;
            }            
        }

        /// <summary>
        /// Get the number of pages in the report
        /// Return -1 if an error occurs
        /// </summary>
        public int PagesCount
        {
            get
            {
                try
                {
                    return GetBitmapDecoder().Frames.Count;
                }
                catch
                {
                    return -1;
                }
            
            }
        }

        /// <summary>
        /// Get a BitmapDecoder of this report 
        /// </summary>
        /// <returns>Return a BitmapDecoder of this current report. Return null on error</returns>
        public BitmapDecoder GetBitmapDecoder()
        {
            try
            {
                if (_dec == null)
                {
                    Stream mStream = new MemoryStream(GetImageArray());
                    _dec = BitmapDecoder.Create(mStream, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.Default);
                    return _dec;
                }
                else
                    return _dec;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Get ou Set the number of copy to print
        /// </summary>
        [DefaultValue(1)]
        public int CopyNumber
        {
            get
            {
                return _Copies;
            }
            set
            {
                _Copies = value;
            }
        }

        /// <summary>
        /// Get or set the path to print rapport file
        /// </summary>
        [DefaultValue("")]
        public string Path
        {
            get
            {
                return _path;
            }
            set
            {
                _path = value;
            }
        }

        /// <summary>
        /// Get ou set the current report
        /// </summary>        
        public LocalReport Report
        {
            get
            {
                return _report;
            }
            set
            {
                _report = value;
            }
        }
        #endregion

        #region Export to EMF (Enhanced Metafile) and stream creation
        /// <summary>
        /// Create a stream user for each raport page...
        /// </summary>
        private Stream CreateStream(string name, string fileNameExtension, Encoding encoding, string mimeType, bool willSeek)
        {
            Stream stream = new MemoryStream();
            _streams.Add(stream);
            return stream;
        }


        // Export to EMF (Enhanced Metafile) format.
        private void Export(LocalReport report)
        {
            if (_printDoc == null)
                _printDoc = GetDefaultPrinter();

            string deviceInfo = GetDeviceInfo();

            Warning[] warnings;
            _streams = new List<Stream>();

            for (int i = 0; i < _Copies; i++)
                report.Render("Image", deviceInfo, CreateStream, out warnings);
            
            foreach (Stream stream in _streams)
                stream.Position = 0;
        }

        /// <summary>
        /// the device information
        /// </summary>
        /// <returns></returns>
        private string GetDeviceInfo(){           

            string deviceinfo;

            if (_report != null)
            {
                if (_report.GetDefaultPageSettings().IsLandscape)
                    deviceinfo = @"<DeviceInfo>
                        <OutputFormat>EMF</OutputFormat>
                        <StartPage>" + _printDoc.PrinterSettings.FromPage.ToString() + @"</StartPage>
                        <EndPage>" + _printDoc.PrinterSettings.ToPage.ToString() + @"</EndPage>
                        <PageWidth>" + ((double)_report.GetDefaultPageSettings().PaperSize.Height / 100).ToString() + @"in</PageWidth>
                        <PageHeight>" + ((double)_report.GetDefaultPageSettings().PaperSize.Width / 100).ToString() + @"in</PageHeight>
                        <MarginTop>" + ((double)_report.GetDefaultPageSettings().Margins.Top / 100).ToString() + @"in</MarginTop>
                        <MarginLeft>" + ((double)_report.GetDefaultPageSettings().Margins.Left / 100).ToString() + @"in</MarginLeft>
                        <MarginRight>" + ((double)_report.GetDefaultPageSettings().Margins.Right / 100).ToString() + @"in</MarginRight>
                        <MarginBottom>" + ((double)_report.GetDefaultPageSettings().Margins.Bottom / 100).ToString() + @"in</MarginBottom>
                    </DeviceInfo>";
                else
                    deviceinfo = @"<DeviceInfo>
                        <OutputFormat>EMF</OutputFormat>
                        <StartPage>" + _printDoc.PrinterSettings.FromPage.ToString() + @"</StartPage>
                        <EndPage>" + _printDoc.PrinterSettings.ToPage.ToString() + @"</EndPage>
                        <PageWidth>" + ((double)_report.GetDefaultPageSettings().PaperSize.Width / 100).ToString() + @"in</PageWidth>
                        <PageHeight>" + ((double)_report.GetDefaultPageSettings().PaperSize.Height / 100).ToString() + @"in</PageHeight>
                        <MarginTop>" + ((double)_report.GetDefaultPageSettings().Margins.Top / 100).ToString() + @"in</MarginTop>
                        <MarginLeft>" + ((double)_report.GetDefaultPageSettings().Margins.Left / 100).ToString() + @"in</MarginLeft>
                        <MarginRight>" + ((double)_report.GetDefaultPageSettings().Margins.Right / 100).ToString() + @"in</MarginRight>
                        <MarginBottom>" + ((double)_report.GetDefaultPageSettings().Margins.Bottom / 100).ToString() + @"in</MarginBottom>
                    </DeviceInfo>";
            }
            else
                deviceinfo = "ERROR";

            return deviceinfo;
        }

        #endregion

        #region Method Print Page
        private void PrintPage(object sender, PrintPageEventArgs ev)
        {
            
            _pageImage = new Metafile(_streams[_currentPageIndex]);

            // Ajuster le rectangle au marge de la page
            _adjustedRect = new System.Drawing.Rectangle(
                ev.PageBounds.Left - (int)ev.PageSettings.HardMarginX,
                ev.PageBounds.Top - (int)ev.PageSettings.HardMarginY,
                ev.PageBounds.Width,
                ev.PageBounds.Height);

            // Dessiner un rectangle blanc sur la page
            ev.Graphics.FillRectangle(Brushes.White, _adjustedRect);

            // Dessiner le rapport
            ev.Graphics.DrawImage(_pageImage, _adjustedRect);
            
            // Prepare la prochaine page.
            _currentPageIndex++;
            ev.HasMorePages = (_currentPageIndex < _streams.Count);
        }
        #endregion

        #region Method Print Now
        private void PrintNow()
        {

            if (_streams == null || _streams.Count == 0)
                return;

                     
            if (!_printDoc.PrinterSettings.IsValid)
            {
                MessageBox.Show("Error: cannot find the default printer", "Print Error" );                
                return;
            }

            _printDoc.PrintPage += new PrintPageEventHandler(PrintPage);
            _printDoc.EndPrint += printDoc_EndPrint;

            _printDoc.Print();
        }
        
        private void printDoc_EndPrint(object sender, PrintEventArgs e)
        {
            _currentPageIndex = 0;
            _streams = null;            
            
            _pageImage = null;
            _adjustedRect = new Rectangle();
            _warnings = null;
            _streamids = null;
            _mimeType = String.Empty;
            _encoding = String.Empty;
            _extension = String.Empty;            
            _printDoc = null;            
        }

        /// <summary>
        /// Get the default printer
        /// </summary>
        /// <returns></returns>
        public PrintDocument GetDefaultPrinter()
        {
            //obtien le nombre de page du rapport
            BitmapDecoder dec = GetImage();

            PrintDocument pDoc = new PrintDocument();
            
            pDoc.DefaultPageSettings.Landscape = _report.GetDefaultPageSettings().IsLandscape;
            pDoc.DefaultPageSettings.PaperSize = new PaperSize(_report.GetDefaultPageSettings().PaperSize.PaperName, _report.GetDefaultPageSettings().PaperSize.Width, _report.GetDefaultPageSettings().PaperSize.Height);
            
            //par default ont imprime toute les pages
            pDoc.PrinterSettings.FromPage = 1;
            pDoc.PrinterSettings.ToPage = dec.Frames.Count;
            return pDoc;
        }


        /// <summary>
        /// Retourne un tableau (Array) d'image sous forme de BitmapDecoder
        /// </summary>
        /// <returns></returns>
        public BitmapDecoder GetImage()
        {            
            Stream mStream = new MemoryStream(GetImageArray());

            return BitmapDecoder.Create(mStream, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.Default);            
        }

        #endregion

        #region method to Print / export to various file format
        /// <summary>
        /// Launch printing
        /// </summary>
        public void Print()
        {
            switch (_ReportType)
            {
                case ReportType.Printer:
                    Export(_report);
                    PrintNow();
                    break;
                case ReportType.Word:
                    SaveAsWord();
                    break;
                case ReportType.PDF:
                    SaveAsPDF();
                    break;
                case ReportType.Excel:
                    SaveAsExcel(); 
                    break;
                case ReportType.Image:
                    SaveAsImage();
                    break;
            }
        }
        #endregion

        #region Some export method


        /// <summary>
        /// Save as PNG image to the specified path
        /// </summary>    
        /// <returns>Return true if the export as completed successfuly</returns>
        public bool SaveAsImage()
        {
            byte[] byteViewer = GetImageArray();

            if (byteViewer != null)
            {
                if (System.IO.Path.GetFileName(_path).Contains(".png"))
                    _filename = _path;
                else
                    _filename = _path + ".png";

                if (FileSaving != null)
                    FileSaving(this, new EventArgs());

                FileStream newFile = new FileStream(_filename, FileMode.Create);
                
                newFile.Write(byteViewer, 0, byteViewer.Length);
                newFile.Close();

                if (FileSaved != null)
                    FileSaved(this, new EventArgs());

                return true;
            }
            else
                return false;
        }


        /// <summary>
        /// Get bitmap image bytes array
        /// </summary>
        /// <returns></returns>
        public byte[] GetImageArray()
        {
            if (_report != null)
                try
                {
                    return _report.Render("Image");
                }
                catch
                {
                    throw new Exception("Error in the render Image method. ");
                }
            else
                return null;
        }


        /// <summary>
        /// Return a BitmapImage that contains report
        /// </summary>
        /// <returns></returns>
        public BitmapImage GetBitmapImage()
        {
            try
            {
                byte[] img = GetImageArray();

                MemoryStream mStream = new MemoryStream(img);

                BitmapImage reportBitmap = new BitmapImage();
                reportBitmap.BeginInit();
                reportBitmap.StreamSource = mStream;
                reportBitmap.EndInit();
                return reportBitmap;
            }
            catch
            {
                return null;
            }

        }
               

        /// <summary>
        /// Save as PDF file to the specified path
        /// </summary>
        /// <returns>Return true if the export as completed successfuly</returns>
        public bool SaveAsPDF()
        {
            byte[] byteViewer = GetPDFArray();

            if (byteViewer != null)
            {
                if (System.IO.Path.GetFileName(_path).Contains(".pdf"))
                    _filename = _path;
                else
                    _filename = _path + ".pdf";

                if (FileSaving != null)
                    FileSaving(this, new EventArgs());

                FileStream newFile = new FileStream(_filename, FileMode.Create);
                newFile.Write(byteViewer, 0, byteViewer.Length);
                newFile.Close();

                if (FileSaved != null)
                    FileSaved(this, new EventArgs());

                return true;
            }
            else
                return false;
        }

        /// <summary>
        /// Get the pdf bytes array
        /// </summary>
        /// <returns></returns>
        public byte[] GetPDFArray()
        {
            if (_report != null)
                return _report.Render("PDF");
            else
                return null;
        }

        /// <summary>
        /// Save as Microsoft Excel file format to the specified path
        /// </summary>
        /// <returns>Return true if the export as completed successfuly</returns>
        public bool SaveAsExcel()
        {
            byte[] bytesViewer = GetExcelArray();
            
            if (bytesViewer != null)  
            {
                if (System.IO.Path.GetFileName(_path).Contains(".xls"))
                    _filename = _path;
                else
                    _filename = _path + ".xls";

                if (FileSaving != null)
                    FileSaving(this, new EventArgs());

                FileStream newFile = new FileStream(_filename, FileMode.Create);
                newFile.Write(bytesViewer, 0, bytesViewer.Length);
                newFile.Close();

                if (FileSaved != null)
                    FileSaved(this, new EventArgs());

                return true;
            }
            else
                return false;
        }

        /// <summary>
        /// Get the Microsoft Excel byte Array
        /// </summary>        
        public byte[] GetExcelArray()
        {
            if (_report != null)
                return _report.Render("Excel", null, out _mimeType, out _encoding,out _extension,out _streamids, out _warnings);
            else
                return null;
        }

        /// <summary>
        /// Save as Microsoft Word file format to the specified path
        /// </summary>
        /// <returns>Return true if the export as completed successfuly</returns>
        public bool SaveAsWord()
        {
            byte[] bytesViewer = GetWordArray();

            if (bytesViewer != null)
            {
                if (System.IO.Path.GetFileName(_path).Contains(".doc"))
                    _filename = _path;
                else
                    _filename = _path + ".doc";

                if (FileSaving != null)
                    FileSaving(this, new EventArgs());
           
                FileStream newFile = new FileStream(_filename, FileMode.Create);
                newFile.Write(bytesViewer, 0, bytesViewer.Length);
                newFile.Close();

                if (FileSaved != null)
                    FileSaved(this, new EventArgs());

                return true;
            }
            else
                return false;
        }

        /// <summary>
        /// Get the Microsoft Word byte Array
        /// </summary>        
        public byte[] GetWordArray()
        {
            if (_report != null)
                return _report.Render("Word", null, out _mimeType, out _encoding, out _extension, out _streamids, out _warnings);
            else
                return null;
        }

        #endregion

        #region Dipose des Streams
        /// <summary>
        /// Dispose all stream
        /// </summary>
        public void Dispose()
        {
            if (_streams != null)
            {
                foreach (Stream stream in _streams)
                    stream.Close();
                _streams = null;
            }
        }
        #endregion
    }
}
