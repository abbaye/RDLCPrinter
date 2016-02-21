using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;

namespace DSoft.MethodExtension
{
    public static class BitmapExtention
    {
        /// <summary>
        /// Recoit un bitmap et retourne la couleur de la premiere pixel du coin supérieur gauche pour etre utiliser pour définir le backgroud d'un control
        /// <remarks>
        /// CREDIT : 2013 Martin Savard
        /// https://github.com/abbaye/RDLCPrinter
        /// </remarks>
        /// </summary>

        public static System.Windows.Media.Color GetColor(this Bitmap b)
        {
            Bitmap myBitmap = new Bitmap(b);

            Color pixelColor = myBitmap.GetPixel(1, 1);

            System.Drawing.Color drawingColor = pixelColor;

            var colorBrush = System.Windows.Media.Color.FromArgb(drawingColor.A, drawingColor.R, drawingColor.G, drawingColor.B);

            return colorBrush;
        }
    }
}
