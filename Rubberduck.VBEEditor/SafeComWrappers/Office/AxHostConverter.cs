using System.Drawing;
using System.Windows.Forms;
using stdole;

namespace Rubberduck.VBEditor.SafeComWrappers
{
    public class AxHostConverter : AxHost
    {
        private AxHostConverter() 
            : base(string.Empty) { }

        public static IPictureDisp ImageToPictureDisp(Image image)
        {
            return (IPictureDisp)GetIPictureDispFromPicture(image);
        }

        public static Image PictureDispToImage(IPictureDisp pictureDisp)
        {
            return GetPictureFromIPicture(pictureDisp);
        }
    }
}