using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;
using System.Windows.Controls;

namespace DSoft.MethodExtension
{
    public static class ImageExtention
    {
        public static BitmapImage GetImageSource(this Image img, string path, UriKind uriKind = UriKind.RelativeOrAbsolute)
        {
            BitmapImage bi = new BitmapImage();

            bi.BeginInit();
            bi.UriSource = new Uri(path, uriKind);
            bi.EndInit();

            return bi;
        }
    }
}
