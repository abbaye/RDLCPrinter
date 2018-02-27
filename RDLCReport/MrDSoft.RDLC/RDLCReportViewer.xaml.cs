using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Linq;
using DSoft.RDLCReport;
using DSoft.MethodExtension;

namespace DSoft.RDLC
{
    /// <summary>
    /// RDLCPreviewControl
    /// <remarks>
    /// CREDIT : 2013-2018 Derek Tremblay (abbaye), 2013 Martin Savard
    /// https://github.com/abbaye/RDLCPrinter
    /// </remarks>
    /// </summary>
    public partial class RDLCReportViewer
    {
        private RDLCPrinter _report;        
        private int _pos;
        private bool _isShowToolBar = true;        
        private Point _origin;
        private Point _start;
        private TranslateTransform tt = new TranslateTransform();
        private bool _fixedToWindowMode;

        public RDLCReportViewer()
        {
            InitializeComponent();

            // zoom initialization
            CreateTransformGroup();
            ZoomSlider.Value = 98;
            ZoomValueTextBloc.Text = (ZoomSlider.Value + 2).ToString("##0");
        }

        private void CreateTransformGroup()
        {
            
            if(ZoomGroup.Children.Count > 2)
                ZoomGroup.Children.Remove(tt);

            tt = new TranslateTransform();
            ZoomGroup.Children.Add(tt);
            

            PreviewImage.RenderTransform = ZoomGroup;

            PreviewImage.MouseLeftButtonDown += PreviewImage_MouseLeftButtonDown;
            PreviewImage.MouseLeftButtonUp += PreviewImage_MouseLeftButtonUp;
            PreviewImage.MouseMove += PreviewImage_MouseMove;
            if(_fixedToWindowMode == false)
                ZoomSlider.Value = 98;
        }

        /// <summary>
        /// For move the report in usercontrol
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void PreviewImage_MouseMove(object sender, MouseEventArgs e)
        {
            if (!PreviewImage.IsMouseCaptured ) 
                return;

            var translateTransf = (TranslateTransform)((TransformGroup)PreviewImage.RenderTransform).Children.First(tr => tr is TranslateTransform);
            var v = _start - e.GetPosition(ImgBorber);
            
            translateTransf.Y = _origin.Y - v.Y;
            translateTransf.X = _origin.X - v.X;
        }

        /// <summary>
        /// Release mouse capture
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void PreviewImage_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            PreviewImage.ReleaseMouseCapture();
        }

        /// <summary>
        /// Call refresh button
        /// </summary>        
        private void cmdRefresh_Click(object sender, RoutedEventArgs e)
        {            
            RefreshControl();
        }

        /// <summary>
        /// Print report button
        /// </summary>        
        private void cmdPrint_Click(object sender, RoutedEventArgs e)
        {
            _report.Reporttype = ReportType.Printer;
            _report.Print();
            
        }

        /// <summary>
        /// Get or Set the RDLC report
        /// </summary>
        public RDLCPrinter Report
        {
            get
            {
                return _report;
            }
            set
            {
                _report = value;
                
                RefreshControl();                

                GiveFocus();
            }
        }

        /// <summary>
        /// Give the focus to usercontrol
        /// </summary>
        public void GiveFocus()
        {
            FocusManager.SetFocusedElement(this, PreviewImage);
            Keyboard.Focus(PreviewImage);
        }


        /// <summary>
        /// Display or not the toolbar
        /// </summary>
        public bool isShowToolBar
        {
            get => _isShowToolBar;
            set
            {
                _isShowToolBar = value;

                ToolBarRow.Visibility = value ? Visibility.Visible : Visibility.Collapsed;
            }
        }

        /// <summary>
        /// Refresh logic.         
        /// </summary>
        public void RefreshControl()
        {
            if (_report == null) return;

            _report.Refresh();

            DispatcherHelper.DoEvents(); //Clear the ui message queud

            LoadImage();
                
            _pos = 0;

            PageSpinner.Maximum = _report.PagesCount;
            PageSpinner.Value = _pos + 1;

            ChangeImage(_pos);

            CreateTransformGroup();
        }

        /// <summary>
        /// Update the toolbar button
        /// </summary>
        private void UpdateToolBarButton()
        {
            if (_report != null)
            {
                TBBRefresh.EnableButton();
                TBBPrint.EnableButton();
                ExportDefault.EnableButton();
                ExportMenu.IsEnabled = true;
                ExportMenu.Opacity = 1;
                TBBPrintWithProperties.EnableButton();
                ZoomInfoStackPanel.Visibility = Visibility.Visible;
                ZoomPopupButton.Visibility = Visibility.Visible;
                ZoomPopupButton.IsEnabled = true;
                ZoomPopupButton.Opacity = 1;
            }
            else
            {
                TBBRefresh.DisableButton();
                TBBPrint.DisableButton();
                ExportDefault.DisableButton();
                ExportMenu.IsEnabled = true;
                ExportMenu.Opacity = 1;
                TBBPrintWithProperties.DisableButton();
                ZoomInfoStackPanel.Visibility = Visibility.Collapsed;
                ZoomPopupButton.Visibility = Visibility.Collapsed;
                ZoomPopupButton.IsEnabled = false;
                ZoomPopupButton.Opacity = 0.5;
            }
            
            if (_pos == 0)
            {
                PagerSeparator.Visibility = Visibility.Visible;
                PreviousImage.DisableButton();
                FirstImage.DisableButton(); 
                NextImage.EnableButton();
                LastImage.EnableButton();
            }
            else
            {
                if (_pos + 1 == _report.PagesCount)
                {
                    NextImage.DisableButton();
                    LastImage.DisableButton();
                    PreviousImage.EnableButton();
                    FirstImage.EnableButton();
                }
                else
                {
                    PreviousImage.EnableButton();
                    NextImage.EnableButton();
                    FirstImage.EnableButton();
                    LastImage.EnableButton();
                }
            }

            if (_report.PagesCount > 1)
            {
                PagerSeparator.Visibility = Visibility.Visible;
                PreviousImage.Visibility = Visibility.Visible;
                NextImage.Visibility = Visibility.Visible;
                FirstImage.Visibility = Visibility.Visible;
                LastImage.Visibility = Visibility.Visible;
                PageSpinner.Visibility = Visibility.Visible;
            }
            else
            {
                PagerSeparator.Visibility = Visibility.Collapsed;
                PreviousImage.Visibility = Visibility.Collapsed;
                NextImage.Visibility = Visibility.Collapsed;
                FirstImage.Visibility = Visibility.Collapsed;
                LastImage.Visibility = Visibility.Collapsed;
                PageSpinner.Visibility = Visibility.Collapsed;
            } 
        }

        /// <summary>
        /// Load image for the preview
        /// </summary>
        private void LoadImage()
        {

            if (_pos == 0)
            {
                UpdateToolBarButton();
                PreviousImage.IsEnabled = false;
                PreviousImage.Opacity = 0.5;
            }
            ChangeImage(_pos);            

        }
        
        /// <summary>
        /// Chage page to the position
        /// </summary>
        /// <param name="position"></param>
        private void ChangeImage(int position)
        {
            var pagecount = _report.GetBitmapDecoder().Frames.Count;

            //Check interval
            if (position < 0)
                position = 0;
            else if (position > pagecount)
                position = pagecount;

            CreateTransformGroup();

            PreviewImage.Source = _report.GetBitmapDecoder().Frames[position];

            UpdateToolBarButton();
        }

        /// <summary>
        /// Set the zoom to "Fit to Window" mode
        /// </summary>
        public void SetFitToWindowMode()
        {
            CreateTransformGroup();

            if (_report == null) return;

            _fixedToWindowMode = true;

            ImageScrollViewer.VerticalScrollBarVisibility = ScrollBarVisibility.Auto;

            PreviewImage.Stretch = Stretch.UniformToFill;

            ZoomSlider.Value = 99;
        }

        private void ExportMethod(ReportType rType)
        {
            var saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();

            switch (rType)
            {
                case ReportType.PDF:
                    saveFileDialog1.Filter = @"Adobe PDF (*.pdf)|*.pdf";
                    saveFileDialog1.FileName = _report.Report.DisplayName + ".pdf";
                    break;
                case ReportType.Excel:
                    saveFileDialog1.Filter = @"Microsoft Excel (*.xls)|*.xls";
                    saveFileDialog1.FileName = _report.Report.DisplayName + ".xls";
                    break;
                case ReportType.Word:
                    saveFileDialog1.Filter = @"Microsoft Word (*.doc)|*.doc";
                    saveFileDialog1.FileName = _report.Report.DisplayName + ".doc";
                    break;
                case ReportType.Image:
                    saveFileDialog1.Filter = @"Image PNG (*.png)|*.png";
                    saveFileDialog1.FileName = _report.Report.DisplayName + ".png";
                    break;
            }

            //saveFileDialog1.FilterIndex = 2;
            saveFileDialog1.RestoreDirectory = true;

            if (saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                _report.Path = saveFileDialog1.FileName;
                _report.Reporttype = rType;
                _report.Print();
            }
        }

        /// <summary>
        /// Previous page
        /// </summary>
        private void PreviousImage_Click(object sender, RoutedEventArgs e)
        {

            Application.Current.MainWindow.Cursor = Cursors.Wait;

            if (_pos > 0)
            {
                _pos--;
                PageSpinner.Value = _pos + 1;
                ChangeImage(_pos);
            }

            Application.Current.MainWindow.Cursor = null;
        }


        /// <summary>
        /// Next page
        /// </summary>
        private void NextImage_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.MainWindow.Cursor = Cursors.Wait;

            if (_pos + 1 < _report.PagesCount)
            {
                _pos++;
                PageSpinner.Value = _pos + 1;
                ChangeImage(_pos);
            }

            Application.Current.MainWindow.Cursor = null;
        }

        /// <summary>
        /// First page
        /// </summary>
        private void FirstImage_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.MainWindow.Cursor = Cursors.Wait;

            _pos = 0;
            ChangeImage(_pos);
            PageSpinner.Value = _pos + 1;

            Application.Current.MainWindow.Cursor = null;
        }

        /// <summary>
        /// Last page
        /// </summary>        
        private void LastImage_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.MainWindow.Cursor = Cursors.Wait;

            _pos = _report.PagesCount - 1;
            PageSpinner.Value = _pos + 1; 
            ChangeImage(_pos);

            Application.Current.MainWindow.Cursor = null;
        }

        /// <summary>
        /// Export to word file format button
        /// </summary>        
        private void MenuItemWord_Click(object sender, RoutedEventArgs e)
        {
            ExportMethod(ReportType.Word);
        }

        /// <summary>
        /// Export to excel file format button
        /// </summary>        
        private void MenuItemExcel_Click(object sender, RoutedEventArgs e)
        {
            ExportMethod(ReportType.Excel);
        }

        /// <summary>
        /// Export to PNG file format button
        /// </summary>        
        private void MenuItemPNG_Click(object sender, RoutedEventArgs e)
        {
            ExportMethod(ReportType.Image);
        }


        /// <summary>
        /// Export to PDF file format button
        /// </summary>        
        private void ExportDefault_Click(object sender, RoutedEventArgs e)
        {
            ExportMethod(ReportType.PDF);
        }

        /// <summary>
        /// Load toolbar logic
        /// </summary>
        private void ToolBarRow_Loaded(object sender, RoutedEventArgs e)
        {
            var toolBar = sender as ToolBar;

            if (toolBar.Template.FindName("OverflowGrid", toolBar) is FrameworkElement overflowGrid)            
                overflowGrid.Visibility = Visibility.Collapsed;
            
            if (toolBar.Template.FindName("MainPanelBorder", toolBar) is FrameworkElement mainPanelBorder)            
                mainPanelBorder.Margin = new Thickness(0);
            
        }

        /// <summary>
        /// Call a RDLCPrintDialog
        /// </summary>
        private void TBBPrintWithProperties_Click(object sender, RoutedEventArgs e)
        {
            var printerDialog = new RDLCPrinterDialog();

            if (_report != null)                            
                printerDialog.Report = _report;            
            else
                return;

            printerDialog.ShowDialog();
        }


        /// <summary>
        /// Drag picture logic
        /// </summary>        
        private void PreviewImage_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (ZoomSlider.Value == 98 || ZoomSlider.Value == 99) return;

            PreviewImage.CaptureMouse();
            var translateTransf = (TranslateTransform)((TransformGroup)PreviewImage.RenderTransform).Children.First(tr => tr is TranslateTransform);
            _start = e.GetPosition(ImgBorber);
            _origin = new Point(translateTransf.X, translateTransf.Y);
        }

        /// <summary>
        /// Zoom slider logic
        /// </summary>        
        private void ZoomSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (ZoomSlider.Value != 98)
            {
                ActualSizeButton.IsEnabled = true;
                ActualSizeButton.Opacity = 1;                
            }
            else
            {
                ActualSizeButton.IsEnabled = false;
                ActualSizeButton.Opacity = 0.5;
            }
            ZoomValueTextBloc.Text = (ZoomSlider.Value + 2).ToString("##0");
        }

        private void ZoomPopupButton_Click(object sender, RoutedEventArgs e) =>
            ZoomPopup.IsOpen = ZoomPopupButton.IsChecked == true;

        #region zoom button event
        private void ActualSizeButton_Click(object sender, RoutedEventArgs e)
        {
            ImageScrollViewer.VerticalScrollBarVisibility = ScrollBarVisibility.Disabled;
            PreviewImage.Stretch = Stretch.Uniform;
            _fixedToWindowMode = false;

            CreateTransformGroup();
            ZoomSlider.Value = 98;
        }

        private void ZoomPopup_Closed(object sender, System.EventArgs e) => ZoomPopupButton.IsChecked = false;


        private void PerCent50Button_Click(object sender, RoutedEventArgs e)
        {
            ZoomSlider.Value = 48;
        }

        private void PerCent100Button_Click(object sender, RoutedEventArgs e)
        {
            ActualSizeButton.PerformClick();
        }

        private void PerCent150Button_Click(object sender, RoutedEventArgs e)
        {
            ZoomSlider.Value = 148;
        }

        private void PerCent200Button_Click(object sender, RoutedEventArgs e)
        {
            ZoomSlider.Value = 198;
        }

        private void PerCent250Button_Click(object sender, RoutedEventArgs e)
        {
            ZoomSlider.Value = 248;
        }
        #endregion //différents click des boutons du zoom


        private void ZoomSlider_MouseWheel(object sender, MouseWheelEventArgs e)
        {
            ZoomSlider.Value += e.Delta / 15;
        }

        private void FitToWindowButton_Click(object sender, RoutedEventArgs e)
        {
            SetFitToWindowMode();        
        }
        
        private void CommandBinding_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            var printerDialog = new RDLCPrinterDialog();

            if (_report != null)                            
                printerDialog.Report = _report;            
            else
                return;

            printerDialog.ShowDialog();
        }        
    }    
}
