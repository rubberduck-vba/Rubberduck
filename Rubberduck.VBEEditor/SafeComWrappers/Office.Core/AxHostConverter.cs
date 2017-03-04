using System.Drawing;
using System.Windows.Forms;
using stdole;

namespace Rubberduck.VBEditor.SafeComWrappers.Office.Core
{
    internal class AxHostConverter : AxHost
    {
        private AxHostConverter() 
            : base(string.Empty) { }

        static public IPictureDisp ImageToPictureDisp(Image image)
        {
            return (IPictureDisp)GetIPictureDispFromPicture(image);
        }

        static public Image PictureDispToImage(IPictureDisp pictureDisp)
        {
            return GetPictureFromIPicture(pictureDisp);
        }
    }
}