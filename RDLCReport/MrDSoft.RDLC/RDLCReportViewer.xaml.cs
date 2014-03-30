using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Linq;
using System.Windows.Media.Effects;
using System;
using DSoft.RDLCReport;


namespace DSoft.RDLC
{
    /// <summary>
    /// Logique d'interaction pour RDLCPreviewControl.xaml
    /// </summary>
    public partial class RDLCReportViewer : UserControl
    {
        private RDLCPrinter _report = null;        
        private byte[] _BitmapArray;
        private int _pos = 0;
        private BitmapDecoder _dec = null;
        private bool _isShowToolBar = true;
        private bool _firstRun = true;
        private Point _origin;
        private Point _start;
        private TranslateTransform tt = new TranslateTransform();
        private bool _fixedToWindowMode = false;
        //private byte[] _PDFFile;
        private Mode _isMode = Mode.RDLC;

        public enum Mode
        { 
            RDLC
            //PDF
        }

        public RDLCReportViewer()
        {
            InitializeComponent();


            // Ajoute les éléments qui permet le zoom et le drag de l,image
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
            if(this._fixedToWindowMode == false)
                ZoomSlider.Value = 98;
        }

        /// <summary>
        /// logique pour bopuger l'image (rapport dans notre cas)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void PreviewImage_MouseMove(object sender, MouseEventArgs e)
        {
            if (!PreviewImage.IsMouseCaptured ) 
                return;

            var tt = (TranslateTransform)((TransformGroup)PreviewImage.RenderTransform).Children.First(tr => tr is TranslateTransform);
            Vector v = _start - e.GetPosition(ImgBorber);
            
            tt.Y = _origin.Y - v.Y;
            tt.X = _origin.X - v.X;
        }

        /// <summary>
        /// relache l'image
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void PreviewImage_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            PreviewImage.ReleaseMouseCapture();
        }

        /// <summary>
        /// Call du refresh au changement de rapport
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cmdRefresh_Click(object sender, RoutedEventArgs e)
        {            
            RefreshControl();
        }

        /// <summary>
        /// Impression du Rapport au clic du bouton Print
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cmdPrint_Click(object sender, RoutedEventArgs e)
        {
            this._report.Reporttype = ReportType.Printer;
            this._report.Print();
            
        }

        ///// <summary>
        ///// Renvoie le fichier PDF demandé
        ///// </summary>
        //public byte[] PDFFile
        //{
        //    get
        //    {
        //        return this._PDFFile;
        //    }
        //    set
        //    {
        //        this._PDFFile = value;

        //        //Logique...
        //        this.isMode = Mode.PDF;
        //        RefreshControl();
        //    }
        //}


        /// <summary>
        /// Renvoie le rpport demandé
        /// </summary>
        public RDLCPrinter Report
        {
            get
            {
                return this._report;
            }
            set
            {
                this._report = value;
                this.isMode = Mode.RDLC;
                RefreshControl();

                GiveFocus();
            }
        }

        public void GiveFocus()
        {
            FocusManager.SetFocusedElement(this, this.PreviewImage);
            Keyboard.Focus(this.PreviewImage);
        }


        public Mode isMode
        {
            get
            {
                return this._isMode;
            }
            set
            {
                this._isMode = value;

                RefreshControl();
            }
        }
        
        /// <summary>
        /// Propriété pour la visibilitée de la toolbar
        /// </summary>
        public bool isShowToolBar
        {
            get
            {
                return this._isShowToolBar;
            }
            set
            {
                this._isShowToolBar = value;

                if (value)
                    ToolBarRow.Visibility = System.Windows.Visibility.Visible;
                else
                    ToolBarRow.Visibility = System.Windows.Visibility.Collapsed;
            }
        }

        /// <summary>
        /// Resfresh du controle au changement de rapport
        /// </summary>
        public void RefreshControl()
        {
            //Logic de refresh
            switch(this._isMode)
            {
                case Mode.RDLC:
                    if (this._report != null)
                    {
                        LoadImage();
                    }
                    _pos = 0;

                    if(_dec != null)
                        NumPager.Text = "1 / " + _dec.Frames.Count;

                    _firstRun = false;

                    CreateTransformGroup();
                    break;
            }
        }

        #region Update des propriétées pour les différents bouton et controles de la toolbar
        /// <summary>
        /// Update des propriétées pour les différents bouton et controles de la toolbar
        /// </summary>
        private void UpdateToolBarButton()
        {
            if (_report != null)
            {
                TBBRefresh.IsEnabled = true;
                TBBRefresh.Opacity = 1;
                TBBPrint.IsEnabled = true;
                TBBPrint.Opacity = 1;
                ExportDefault.IsEnabled = true;
                ExportDefault.Opacity = 1;
                ExportMenu.IsEnabled = true;
                ExportMenu.Opacity = 1;
                TBBPrintWithProperties.IsEnabled = true;
                TBBPrintWithProperties.Opacity = 1;
                ZoomInfoStackPanel.Visibility = Visibility.Visible;
                ZoomPopupButton.Visibility = Visibility.Visible;
                ZoomPopupButton.IsEnabled = true;
                ZoomPopupButton.Opacity = 1;
            }
            else
            {
                TBBRefresh.IsEnabled = false;
                TBBRefresh.Opacity = 0.5;
                TBBPrint.IsEnabled = false;
                TBBPrint.Opacity = 0.5;
                ExportDefault.IsEnabled = false;
                ExportDefault.Opacity = 0.5;
                ExportMenu.IsEnabled = true;
                ExportMenu.Opacity = 1;
                TBBPrintWithProperties.IsEnabled = false;
                TBBPrintWithProperties.Opacity = 0.5;
                ZoomInfoStackPanel.Visibility = Visibility.Hidden;
                ZoomPopupButton.Visibility = Visibility.Hidden;
                ZoomPopupButton.IsEnabled = false;
                ZoomPopupButton.Opacity = 0.5;
            }


            if (_pos == 0)
            {
                PagerSeparator.Visibility = Visibility.Visible;
                PreviousImage.IsEnabled = false;
                PreviousImage.Opacity = 0.5;
                FirstImage.IsEnabled = false;
                FirstImage.Opacity = 0.5;
                SpinnerDown.IsEnabled = false;
                SpinnerDown.Opacity = 0.5;
                NextImage.IsEnabled = true;
                NextImage.Opacity = 1;
                LastImage.IsEnabled = true;
                LastImage.Opacity = 1;
                SpinnerUp.IsEnabled = true;
                SpinnerUp.Opacity = 1;
            }
            else
            {
                if (_pos + 1 == _dec.Frames.Count)
                {
                    NextImage.IsEnabled = false;
                    NextImage.Opacity = 0.5;
                    LastImage.IsEnabled = false;
                    LastImage.Opacity = 0.5;
                    SpinnerUp.IsEnabled = false;
                    SpinnerUp.Opacity = 0.5;
                    PreviousImage.IsEnabled = true;
                    PreviousImage.Opacity = 1;
                    FirstImage.IsEnabled = true;
                    FirstImage.Opacity = 1;
                    SpinnerDown.IsEnabled = true;
                    SpinnerDown.Opacity = 1;
                }
                else
                {
                    PreviousImage.IsEnabled = true;
                    PreviousImage.Opacity = 1;
                    NextImage.IsEnabled = true;
                    NextImage.Opacity = 1;
                    FirstImage.IsEnabled = true;
                    FirstImage.Opacity = 1;
                    LastImage.IsEnabled = true;
                    LastImage.Opacity = 1;
                    SpinnerUp.IsEnabled = true;
                    SpinnerUp.Opacity = 1;
                    SpinnerDown.IsEnabled = true;
                    SpinnerDown.Opacity = 1;
                }
            }


            if (_dec.Frames.Count > 1)
            {
                PagerSeparator.Visibility = Visibility.Visible;
                PreviousImage.Visibility = Visibility.Visible;
                NextImage.Visibility = Visibility.Visible;
                NumPager.Visibility = Visibility.Visible;
                FirstImage.Visibility = Visibility.Visible;
                LastImage.Visibility = Visibility.Visible;
                SpinnerGrid.Visibility = Visibility.Visible;
            }
            else
            {
                PagerSeparator.Visibility = Visibility.Hidden;
                PreviousImage.Visibility = Visibility.Hidden;
                NextImage.Visibility = Visibility.Hidden;
                NumPager.Visibility = Visibility.Hidden;
                FirstImage.Visibility = Visibility.Hidden;
                LastImage.Visibility = Visibility.Hidden;
                SpinnerGrid.Visibility = Visibility.Hidden;
            } 
        }
        #endregion

        /// <summary>
        /// Load de l'image en bitmap image pour afficher dans le controle utilisateur
        /// </summary>
        private void LoadImage()
        {
            switch(this.isMode)
            {
                case Mode.RDLC:
                    if(_pos ==0)
                    {
                        this._BitmapArray = _report.GetImageArray();
                        Stream mStream = new MemoryStream(_BitmapArray);
            
                        _dec = BitmapDecoder.Create(mStream, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.Default);
                                
                        UpdateToolBarButton();
                        PreviousImage.IsEnabled = false;
                        PreviousImage.Opacity = 0.5;
                    }
                    ChangeImage(_pos);
                    break;
            }
        }


        /// <summary>
        /// change l'image (page) sur demande
        /// </summary>
        /// <param name="position"></param>
        private void ChangeImage(int position)
        {
            
            CreateTransformGroup();

            this.PreviewImage.Source = _dec.Frames[position];
            
        }

        /// <summary>
        /// Clic du bouton PreviousImage (page précédente)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void PreviousImage_Click(object sender, RoutedEventArgs e)
        {

            Application.Current.MainWindow.Cursor = Cursors.Wait;

            if (_pos > 0)
            {
                _pos--;
                NumPager.Text = (_pos + 1).ToString() + " / " + _dec.Frames.Count;
                ChangeImage(_pos);
            }


            Application.Current.MainWindow.Cursor = null;
        }


        /// <summary>
        /// Clic du bouton NextImage ( page suivante)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void NextImage_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.MainWindow.Cursor = Cursors.Wait;

            if (_pos + 1 < _dec.Frames.Count)
            {
                _pos++;
                NumPager.Text = (_pos + 1).ToString() + " / " + _dec.Frames.Count;
                ChangeImage(_pos);
            }
            Application.Current.MainWindow.Cursor = null;
        }

        /// <summary>
        /// gère les entrées au clavier pour le Integer Up Down (num up down)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void NumPager_KeyDown(object sender, KeyEventArgs e)
        {
            e.Handled = true;
        }



        private void NumPager_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_firstRun != true)
            {
                //_pos = Convert.ToInt32(NumPager.Text) - 1;
                ChangeImage(_pos);
                UpdateToolBarButton();
            }  
        }


        private void FirstImage_Click(object sender, RoutedEventArgs e)
        {
            _pos = 0;
            ChangeImage(_pos);
            NumPager.Text = (_pos + 1).ToString() + " / " + _dec.Frames.Count;
        }

        private void LastImage_Click(object sender, RoutedEventArgs e)
        {
            _pos = _dec.Frames.Count - 1;
            NumPager.Text = (_pos + 1).ToString() + " / " + _dec.Frames.Count;
            ChangeImage(_pos);
        }

        private void SpinnerUp_Click(object sender, RoutedEventArgs e)
        {
            NextImage.PerformClick();
        }

        private void SpinnerDown_Click(object sender, RoutedEventArgs e)
        {
            PreviousImage.PerformClick();
        }


        // Différente logique pour l'importation des différent type de fichier

        private void MenuItemWord_Click(object sender, RoutedEventArgs e)
        {
            ExportMethod(ReportType.Word);
        }

        private void MenuItemExcel_Click(object sender, RoutedEventArgs e)
        {
            ExportMethod(ReportType.Excel);
        }

        private void MenuItemPNG_Click(object sender, RoutedEventArgs e)
        {
            ExportMethod(ReportType.Image);
        }

        private void ExportMethod(ReportType rType)
        {
            System.Windows.Forms.SaveFileDialog saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();

            switch (rType)
            {
                case ReportType.PDF:
                    saveFileDialog1.Filter = "Adobe PDF (*.pdf)|*.pdf";
                    saveFileDialog1.FileName = _report.Report.DisplayName + ".pdf";
                    break;
                case ReportType.Excel:
                    saveFileDialog1.Filter = "Microsoft Excel (*.xls)|*.xls";
                    saveFileDialog1.FileName = _report.Report.DisplayName + ".xls";
                    break;
                case ReportType.Word:
                    saveFileDialog1.Filter = "Microsoft Word (*.doc)|*.doc";
                    saveFileDialog1.FileName = _report.Report.DisplayName + ".doc";
                    break;
                case ReportType.Image:
                    saveFileDialog1.Filter = "Image PNG (*.png)|*.png";
                    saveFileDialog1.FileName = _report.Report.DisplayName + ".png";
                    break;
            }

            //saveFileDialog1.FilterIndex = 2;
            saveFileDialog1.RestoreDirectory = true;

            if (saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                this._report.Path = saveFileDialog1.FileName;
                this._report.Reporttype = rType;
                this._report.Print();
            }
        }

        private void ExportDefault_Click(object sender, RoutedEventArgs e)
        {
            ExportMethod(ReportType.PDF);
        }

        /// <summary>
        /// Logique qui load la ToolBar
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ToolBarRow_Loaded(object sender, RoutedEventArgs e)
        {
            ToolBar toolBar = sender as ToolBar;
            var overflowGrid = toolBar.Template.FindName("OverflowGrid", toolBar) as FrameworkElement;
            if (overflowGrid != null)
            {
                overflowGrid.Visibility = Visibility.Collapsed;
            }

            var mainPanelBorder = toolBar.Template.FindName("MainPanelBorder", toolBar) as FrameworkElement;
            if (mainPanelBorder != null)
            {
                mainPanelBorder.Margin = new Thickness(0);
            }
        }

        /// <summary>
        /// Appel de la fenêtre RDLCPrintDialog
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TBBPrintWithProperties_Click(object sender, RoutedEventArgs e)
        {
            RDLCPrinterDialog printerDialog = new RDLCPrinterDialog();

            switch(this._isMode)
            {
                case Mode.RDLC:
                    if (this._report != null)
                    {
                        printerDialog.NbrPageRapport = _dec.Frames.Count;
                        printerDialog.CurrentReport = this._report;
                    }
                    else
                        return;
                    break;
            }
            printerDialog.ShowDialog();
        }


        /// <summary>
        /// Logique pour le "drag" de l'image
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void PreviewImage_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (ZoomSlider.Value == 98 || ZoomSlider.Value == 99) return;

            PreviewImage.CaptureMouse();
            var tt = (TranslateTransform)((TransformGroup)PreviewImage.RenderTransform).Children.First(tr => tr is TranslateTransform);
            _start = e.GetPosition(ImgBorber);
            _origin = new Point(tt.X, tt.Y);
        }

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

        private void ZoomPopupButton_Click(object sender, RoutedEventArgs e)
        {
            if (ZoomPopupButton.IsChecked == true)
                ZoomPopup.IsOpen = true;
            else
                ZoomPopup.IsOpen = false;
        }

        #region différents click des boutons du zoom
        // Différent click des bouton du Zoom
        private void ActualSizeButton_Click(object sender, RoutedEventArgs e)
        {
            ImageScrollViewer.VerticalScrollBarVisibility = ScrollBarVisibility.Disabled;
            PreviewImage.Stretch = Stretch.Uniform;
            this._fixedToWindowMode = false;

            CreateTransformGroup();
            ZoomSlider.Value = 98;
        }

        private void ZoomPopup_Closed(object sender, System.EventArgs e)
        {
            ZoomPopupButton.IsChecked = false;
        }


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


        /// <summary>
        /// entre en mode Fixé a la fenêtre
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void FitToWindowButton_Click(object sender, RoutedEventArgs e)
        {
            
            CreateTransformGroup();

            if(this._report != null)
            {
                _fixedToWindowMode = true;

                

                ImageScrollViewer.VerticalScrollBarVisibility = ScrollBarVisibility.Auto;

                PreviewImage.Stretch = Stretch.UniformToFill;

                this.ZoomSlider.Value = 99;
            }            
        }

        private void CommandBinding_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            RDLCPrinterDialog printerDialog = new RDLCPrinterDialog();

            switch (this._isMode)
            {
                case Mode.RDLC:
                    if (this._report != null)
                    {
                        printerDialog.NbrPageRapport = _dec.Frames.Count;
                        printerDialog.CurrentReport = this._report;
                    }
                    else
                        return;
                    break;
            }
            printerDialog.ShowDialog();
        }
    }

    /// <summary>
    /// Classe d'extention
    /// </summary>
    public static class BoutonClick
    {
        public static void PerformClick(this System.Windows.Controls.Button btn)
        {
            btn.RaiseEvent(new RoutedEventArgs(System.Windows.Controls.Button.ClickEvent));
        }
    }
}
