//////////////////////////////////////////////
// MIT - 2012-2020
// Author : Derek Tremblay (derektremblay666@gmail.com)
// Contributor : Martin Savard (2013)
// https://github.com/abbaye/RDLCPrinter
//////////////////////////////////////////////

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
    public class RDLCPrinter : IDisposable
    {
        private int _currentPageIndex;
        private IList<Stream> _streams;
        private Metafile _pageImage;
        private Rectangle _adjustedRect;
        private Warning[] _warnings;
        private string[] _streamids;
        private string _mimeType = string.Empty;
        private string _encoding = string.Empty;
        private string _extension = string.Empty;
        private string _filename = "";
        private PrintDocument _printDoc;
        private BitmapDecoder _dec;

        //Event
        public event EventHandler FileSaving;
        public event EventHandler FileSaved;
        public event EventHandler BeforeRefresh;

        /// <summary>
        /// Initialize an empty report object
        /// </summary>
        public RDLCPrinter() { }

        /// <summary>
        /// Initialize report object with default setting
        /// </summary>
        public RDLCPrinter(LocalReport report, ReportType rtype, int nbrPage, string path)
        {
            Report = report;
            CopyNumber = nbrPage;
            Reporttype = rtype;
            Path = path;
        }

        /// <summary>
        /// Initialize repot object with default setting
        /// Copies = 1
        /// ReportType = Printer
        /// Path = 0
        /// </summary>
        public RDLCPrinter(LocalReport report)
        {
            Report = report;
            CopyNumber = 1;
            Reporttype = ReportType.Printer;
            Path = "";
        }

        /// <summary>
        /// Refresh Report and Data
        /// </summary>
        public void Refresh()
        {
            if (Report == null) return;
            
            BeforeRefresh?.Invoke(this, new EventArgs());

            Report.Refresh();
            _dec = null;
        }

        #region Properties

        /// <summary>
        ///Get or set the printer for print report (if null, get default printer)
        /// </summary>        
        public PrintDocument PrintDoc
        {
            get => _printDoc ?? GetDefaultPrinter();
            set => _printDoc = value;
        }
        
        /// <summary>
        /// get or set the type of report
        /// </summary>
        [DefaultValue(ReportType.Printer)]
        public ReportType Reporttype { get; set; }

        /// <summary>
        /// Get the default orientation of report
        /// </summary>
        public bool? IsDefaultLandscape => Report?.GetDefaultPageSettings().IsLandscape;

        /// <summary>
        /// Get the number of pages in the report
        /// Return -1 if an error occurs
        /// </summary>
        public int PagesCount => GetBitmapDecoder()?.Frames.Count ?? -1;

        /// <summary>
        /// Get a BitmapDecoder of this report 
        /// </summary>
        /// <returns>Return a BitmapDecoder of this current report. Return null on error</returns>
        public BitmapDecoder GetBitmapDecoder()
        {
            try
            {
                if (_dec != null) return _dec;

                Stream mStream = new MemoryStream(GetImageArray());
                _dec = BitmapDecoder.Create(mStream, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.Default);
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
            _streams.Add(stream);
            return stream;
        }


        // Export to EMF (Enhanced Metafile) format.
        private void Export(LocalReport report)
        {
            if (_printDoc == null)
                _printDoc = GetDefaultPrinter();

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
            string deviceinfo;

            if (Report != null)
            {
                if (Report.GetDefaultPageSettings().IsLandscape)
                    deviceinfo = $@"<DeviceInfo>
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
                    deviceinfo = $@"<DeviceInfo>
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
            }
            else
                deviceinfo = "ERROR";

            return deviceinfo;
        }

        #endregion

        #region Method Print Page

        private void PrintDoc_PrintPage(object sender, PrintPageEventArgs ev)
        {
            _pageImage = new Metafile(_streams[_currentPageIndex]);

            // Ajuster le rectangle au marge de la page
            _adjustedRect = new Rectangle(
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
            ev.HasMorePages = _currentPageIndex < _streams.Count;
        }

        #endregion

        #region Method Print Now

        private void PrintNow()
        {
            if (_streams == null || _streams.Count == 0)
                return;

            if (!_printDoc.PrinterSettings.IsValid)
            {
                MessageBox.Show(@"Error: cannot find the default printer", @"Print Error");
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
            _warnings = null;
            _streamids = null;
            _mimeType = string.Empty;
            _encoding = string.Empty;
            _extension = string.Empty;
            _printDoc = null;
        }

        /// <summary>
        /// Get the default printer
        /// </summary>
        /// <returns></returns>
        public PrintDocument GetDefaultPrinter() => new PrintDocument
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
                ToPage = GetImage().Frames.Count
            }
        };


        /// <summary>
        /// Retourne un tableau (Array) d'image sous forme de BitmapDecoder
        /// </summary>
        public BitmapDecoder GetImage()
        {
            Stream mStream = new MemoryStream(GetImageArray());

            return BitmapDecoder.Create(mStream, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.Default);
        }

        #endregion

        #region Method to Print / export to various file format

        /// <summary>
        /// Launch printing
        /// </summary>
        public void Print()
        {
            switch (Reporttype)
            {
                case ReportType.Printer:
                    Export(Report);
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
            var byteViewer = GetImageArray();

            if (byteViewer == null) return false;

            _filename = System.IO.Path.GetFileName(Path).Contains(".png") ? Path : Path + ".png";

            FileSaving?.Invoke(this, new EventArgs());

            var newFile = new FileStream(_filename, FileMode.Create);

            newFile.Write(byteViewer, 0, byteViewer.Length);
            newFile.Close();

            FileSaved?.Invoke(this, new EventArgs());

            return true;
        }

        /// <summary>
        /// Get bitmap image bytes array
        /// </summary>
        public byte[] GetImageArray()
        {
            if (Report == null) return null;

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
        /// Return a BitmapImage that contains report
        /// </summary>
        public BitmapImage GetBitmapImage()
        {
            try
            {
                var img = GetImageArray();

                var mStream = new MemoryStream(img);

                var reportBitmap = new BitmapImage();
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
        public void SaveAsPDF()
        {
            var byteViewer = GetPDFArray();

            if (byteViewer == null) return;

            _filename = System.IO.Path.GetFileName(Path).Contains(".pdf") ? Path : Path + ".pdf";

            FileSaving?.Invoke(this, new EventArgs());

            var newFile = new FileStream(_filename, FileMode.Create);
            newFile.Write(byteViewer, 0, byteViewer.Length);
            newFile.Close();

            FileSaved?.Invoke(this, new EventArgs());
        }

        /// <summary>
        /// Get the pdf bytes array
        /// </summary>
        public byte[] GetPDFArray() => Report?.Render("PDF");

        /// <summary>
        /// Save as Microsoft Excel file format to the specified path
        /// </summary>
        /// <returns>Return true if the export as completed successfuly</returns>
        public void SaveAsExcel()
        {
            var bytesViewer = GetExcelArray();

            if (bytesViewer == null) return;

            _filename = System.IO.Path.GetFileName(Path).Contains(".xls") ? Path : Path + ".xls";

            FileSaving?.Invoke(this, new EventArgs());

            var newFile = new FileStream(_filename, FileMode.Create);
            newFile.Write(bytesViewer, 0, bytesViewer.Length);
            newFile.Close();

            FileSaved?.Invoke(this, new EventArgs());
        }

        /// <summary>
        /// Get the Microsoft Excel byte Array
        /// </summary>        
        public byte[] GetExcelArray() =>
            Report?.Render("Excel", null, out _mimeType, out _encoding, out _extension, out _streamids, out _warnings);

        /// <summary>
        /// Save as Microsoft Word file format to the specified path
        /// </summary>
        /// <returns>Return true if the export as completed successfuly</returns>
        public void SaveAsWord()
        {
            var bytesViewer = GetWordArray();

            if (bytesViewer == null) return;

            _filename = System.IO.Path.GetFileName(Path).Contains(".doc") ? Path : Path + ".doc";

            FileSaving?.Invoke(this, new EventArgs());

            var newFile = new FileStream(_filename, FileMode.Create);
            newFile.Write(bytesViewer, 0, bytesViewer.Length);
            newFile.Close();

            FileSaved?.Invoke(this, new EventArgs());
        }

        /// <summary>
        /// Get the Microsoft Word byte Array
        /// </summary>        
        public byte[] GetWordArray() =>
            Report?.Render("Word", null, out _mimeType, out _encoding, out _extension, out _streamids, out _warnings);

        #endregion

        #region IDisposable Support
        private bool disposedValue = false; // for detect redondant call

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    if (_streams == null) return;

                    foreach (var stream in _streams)
                        stream.Close();

                    _pageImage.Dispose();
                    _printDoc.Dispose();

                    _streams = null;
                }

                disposedValue = true;
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
