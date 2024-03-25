//////////////////////////////////////////////
// MIT - 2012-2020
// Author : Derek Tremblay (derektremblay666@gmail.com)
// Contributor : Martin Savard (2013)
// Contributor : Gerhard Stephan (gerhard.stephan@gmail.com)
// https://github.com/abbaye/RDLCPrinter
//////////////////////////////////////////////

using System.ComponentModel;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using System.Text;
using Windows.Graphics.Imaging;
using Microsoft.Reporting.NETCore;

namespace TimePunch.Rdlc
{
    /// <summary>
    /// This class allow to Print and export RDLC Report. 
    /// </summary>
    public class RdlcPrinter : IDisposable
    {
        private int _currentPageIndex;
        private IList<Stream>? _streams;
        private Rectangle _adjustedRect;
        private string _filename = "";

        private PrintDocument? _printDoc = null;
        private BitmapDecoder? _dec = null;
        private Metafile? _pageImage = null;

        //Event
        public event EventHandler? FileSaving;
        public event EventHandler? FileSaved;
        public event EventHandler? BeforeRefresh;

        /// <summary>
        /// Initialize report object with default setting
        /// </summary>
        public RdlcPrinter(LocalReport report, ReportType rtype, int nbrPage, string path)
        {
            Report = report;
            CopyNumber = nbrPage;
            ReportType = rtype;
            Path = path;
        }

        /// <summary>
        /// Initialize repot object with default setting
        /// Copies = 1
        /// ReportType = Printer
        /// Path = 0
        /// </summary>
        public RdlcPrinter(LocalReport report)
        {
            Report = report;
            CopyNumber = 1;
            ReportType = ReportType.Printer;
            Path = "";
        }

        /// <summary>
        /// Refresh Report and Data
        /// </summary>
        public void Refresh()
        {
            BeforeRefresh?.Invoke(this, EventArgs.Empty);

            Report.Refresh();
            _dec = null;
        }

        #region Properties

        /// <summary>
        ///Get or set the printer for print report (if null, get default printer)
        /// </summary>        
        public async Task<PrintDocument?> GetPrintDocAsync()
        {
            if (_printDoc != null)
                return _printDoc;

            return await GetDefaultPrinter();
        }

        public void SetPrintDoc(PrintDocument? value)
        {
            _printDoc = value;
        }
        
        /// <summary>
        /// get or set the type of report
        /// </summary>
        [DefaultValue(ReportType.Printer)]
        public ReportType ReportType { get; set; }

        /// <summary>
        /// Get the default orientation of report
        /// </summary>
        public bool? IsDefaultLandscape => Report?.GetDefaultPageSettings().IsLandscape;


        public async Task<int> GetPagesCountAsync()
        {
            var frameCount = (await GetBitmapDecoderAsync())?.FrameCount;
            if (frameCount != null)
                return (int)frameCount;

            return -1;
        }

        /// <summary>
        /// Get a BitmapDecoder of this report 
        /// </summary>
        /// <returns>Return a BitmapDecoder of this current report. Return null on error</returns>
        public async Task<BitmapDecoder?> GetBitmapDecoderAsync()
        {
            try
            {
                if (_dec != null) 
                    return _dec;

                var mStream = new MemoryStream(GetImageArray()).AsRandomAccessStream();
                _dec = await BitmapDecoder.CreateAsync(mStream); //, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.Default);
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
        public int CopyNumber { get; set; }

        /// <summary>
        /// Get or set the path to print rapport file
        /// </summary>
        [DefaultValue("")]
        public string Path { get; set; }

        /// <summary>
        /// Get ou set the current report
        /// </summary>        
        public LocalReport Report { get; set; }

        #endregion

        #region Export to EMF (Enhanced Metafile) and stream creation

        /// <summary>
        /// Create a stream user for each raport page...
        /// </summary>
        private Stream CreateStream(string name, string fileNameExtension, Encoding encoding, string mimeType, bool willSeek)
        {
            Stream stream = new MemoryStream();
            _streams?.Add(stream);
            return stream;
        }


        // Export to EMF (Enhanced Metafile) format.
        private async Task Export(LocalReport report)
        {
            _printDoc ??= await GetDefaultPrinter();

            _streams = new List<Stream>();

            for (var i = 0; i < CopyNumber; i++)
                report.Render("Image", GetDeviceInfo(), CreateStream, out _);

            foreach (var stream in _streams)
                stream.Position = 0;
        }

        /// <summary>
        /// the device information
        /// </summary>
        /// <returns></returns>
        private string GetDeviceInfo()
        {
            string deviceInfo;

            if (Report.GetDefaultPageSettings().IsLandscape)
                deviceInfo = $@"<DeviceInfo>
                        <OutputFormat>EMF</OutputFormat>
                        <StartPage>{_printDoc.PrinterSettings.FromPage}</StartPage>
                        <EndPage>{_printDoc.PrinterSettings.ToPage}</EndPage>
                        <PageWidth>{(double)Report.GetDefaultPageSettings().PaperSize.Height / 100}in</PageWidth>
                        <PageHeight>{(double)Report.GetDefaultPageSettings().PaperSize.Width / 100}in</PageHeight>
                        <MarginTop>{(double)Report.GetDefaultPageSettings().Margins.Top / 100}in</MarginTop>
                        <MarginLeft>{(double)Report.GetDefaultPageSettings().Margins.Left / 100}in</MarginLeft>
                        <MarginRight>{(double)Report.GetDefaultPageSettings().Margins.Right / 100}in</MarginRight>
                        <MarginBottom>{(double)Report.GetDefaultPageSettings().Margins.Bottom / 100}in</MarginBottom>
                    </DeviceInfo>";
            else
                deviceInfo = $@"<DeviceInfo>
                        <OutputFormat>EMF</OutputFormat>
                        <StartPage>{_printDoc.PrinterSettings.FromPage}</StartPage>
                        <EndPage>{_printDoc.PrinterSettings.ToPage}</EndPage>
                        <PageWidth>{(double)Report.GetDefaultPageSettings().PaperSize.Width / 100}in</PageWidth>
                        <PageHeight>{(double)Report.GetDefaultPageSettings().PaperSize.Height / 100}in</PageHeight>
                        <MarginTop>{(double)Report.GetDefaultPageSettings().Margins.Top / 100}in</MarginTop>
                        <MarginLeft>{(double)Report.GetDefaultPageSettings().Margins.Left / 100}in</MarginLeft>
                        <MarginRight>{(double)Report.GetDefaultPageSettings().Margins.Right / 100}in</MarginRight>
                        <MarginBottom>{(double)Report.GetDefaultPageSettings().Margins.Bottom / 100}in</MarginBottom>
                    </DeviceInfo>";

            return deviceInfo;
        }

        #endregion

        #region Method Print Page

        private void PrintDoc_PrintPage(object sender, PrintPageEventArgs ev)
        {
            if (_streams == null)
                return;

            _pageImage = new Metafile(_streams[_currentPageIndex]);

            // Ajuster le rectangle au marge de la page
            _adjustedRect = ev.PageBounds with { X = ev.PageBounds.Left - (int)ev.PageSettings.HardMarginX, Y = ev.PageBounds.Top - (int)ev.PageSettings.HardMarginY };

            // Dessiner un rectangle blanc sur la page
            ev.Graphics.FillRectangle(Brushes.White, _adjustedRect);

            // Dessiner le rapport
            ev.Graphics.DrawImage(_pageImage, _adjustedRect);

            // Prepare la prochaine page.
            _currentPageIndex++;
            ev.HasMorePages = _currentPageIndex < _streams.Count;
        }

        #endregion

        #region Method Print Now

        private void PrintNow()
        {
            if (_streams == null || _streams.Count == 0 || _printDoc == null)
                return;

            if (!_printDoc.PrinterSettings.IsValid)
            {
                // Todo: Show a adequate message to the user
                //MessageBox.Show(@"Error: cannot find the default printer", @"Print Error");
                return;
            }

            _printDoc.PrintPage += PrintDoc_PrintPage;
            _printDoc.EndPrint += PrintDoc_EndPrint;

            _printDoc.Print();
        }

        private void PrintDoc_EndPrint(object sender, PrintEventArgs e)
        {
            _currentPageIndex = 0;
            _streams = null;

            _pageImage = null;
            _adjustedRect = new Rectangle();
            //_warnings = null;
            //_streamids = null;
            //_mimeType = string.Empty;
            //_encoding = string.Empty;
            //_extension = string.Empty;
            _printDoc = null;
        }

        /// <summary>
        /// Get the default printer
        /// </summary>
        /// <returns></returns>
        public async Task<PrintDocument> GetDefaultPrinter() => new PrintDocument
        {
            DefaultPageSettings =
            {
                Landscape = Report.GetDefaultPageSettings().IsLandscape,
                PaperSize = new PaperSize(Report.GetDefaultPageSettings().PaperSize.PaperName,
                    Report.GetDefaultPageSettings().PaperSize.Width,
                    Report.GetDefaultPageSettings().PaperSize.Height)
            },

            PrinterSettings =
            {
                FromPage = 1,
                ToPage = await GetPagesCountAsync()
            }
        };

        #endregion

        #region Method to Print / export to various file format

        /// <summary>
        /// Launch printing
        /// </summary>
        public async Task Print()
        {
            switch (ReportType)
            {
                case ReportType.Printer:
                    await Export(Report);
                    PrintNow();
                    break;
                case ReportType.Word:
                    SaveAsWord();
                    break;
                case ReportType.Pdf:
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
            var byteViewer = GetImageArray();
            _filename = System.IO.Path.GetFileName(Path).Contains(".png") ? Path : Path + ".png";

            FileSaving?.Invoke(this, EventArgs.Empty);

            var newFile = new FileStream(_filename, FileMode.Create);

            newFile.Write(byteViewer, 0, byteViewer.Length);
            newFile.Close();

            FileSaved?.Invoke(this, EventArgs.Empty);

            return true;
        }

        /// <summary>
        /// Get bitmap image bytes array
        /// </summary>
        public byte[] GetImageArray()
        {
            try
            {
                return Report.Render("Image");
            }
            catch
            {
                throw new Exception("Error in the render Image method. ");
            }
        }

        /// <summary>
        /// Save as PDF file to the specified path
        /// </summary>
        /// <returns>Return true if the export as completed successfully</returns>
        public void SaveAsPDF()
        {
            var byteViewer = GetPDFArray();
            _filename = System.IO.Path.GetFileName(Path).Contains(".pdf") ? Path : Path + ".pdf";

            FileSaving?.Invoke(this, EventArgs.Empty);

            var newFile = new FileStream(_filename, FileMode.Create);
            newFile.Write(byteViewer, 0, byteViewer.Length);
            newFile.Close();

            FileSaved?.Invoke(this, EventArgs.Empty);
        }

        /// <summary>
        /// Get the pdf bytes array
        /// </summary>
        public byte[] GetPDFArray() => Report.Render("PDF");

        /// <summary>
        /// Save as Microsoft Excel file format to the specified path
        /// </summary>
        /// <returns>Return true if the export as completed successfuly</returns>
        public void SaveAsExcel()
        {
            var bytesViewer = GetExcelArray();
            _filename = System.IO.Path.GetFileName(Path).Contains(".xls") ? Path : Path + ".xls";

            FileSaving?.Invoke(this, EventArgs.Empty);

            var newFile = new FileStream(_filename, FileMode.Create);
            newFile.Write(bytesViewer, 0, bytesViewer.Length);
            newFile.Close();

            FileSaved?.Invoke(this, EventArgs.Empty);
        }

        /// <summary>
        /// Get the Microsoft Excel byte Array
        /// </summary>        
        public byte[] GetExcelArray() => Report.Render("Excel", null, out _, out _, out _, out _, out _);

        /// <summary>
        /// Save as Microsoft Word file format to the specified path
        /// </summary>
        /// <returns>Return true if the export as completed successfuly</returns>
        public void SaveAsWord()
        {
            var bytesViewer = GetWordArray();
            _filename = System.IO.Path.GetFileName(Path).Contains(".doc") ? Path : Path + ".doc";

            FileSaving?.Invoke(this, EventArgs.Empty);

            var newFile = new FileStream(_filename, FileMode.Create);
            newFile.Write(bytesViewer, 0, bytesViewer.Length);
            newFile.Close();

            FileSaved?.Invoke(this, EventArgs.Empty);
        }

        /// <summary>
        /// Get the Microsoft Word byte Array
        /// </summary>        
        public byte[] GetWordArray() =>  Report.Render("Word", null, out _, out _, out _, out _, out _);

        #endregion

        #region IDisposable Support
        private bool _disposedValue = false; // for detect redundant call

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposedValue)
            {
                if (disposing)
                {
                    if (_streams == null) return;

                    foreach (var stream in _streams)
                        stream.Close();

                    _pageImage.Dispose();
                    _printDoc?.Dispose();

                    _streams = null;
                }

                _disposedValue = true;
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        #endregion
    }
}
