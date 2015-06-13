using System;
using System.Drawing;
using System.Windows.Forms;

using NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;
using NetOffice.VBIDEApi;
using stdole;
using CommandBarButtonClickEvent = NetOffice.OfficeApi.CommandBarButton_ClickEventHandler;

namespace Rubberduck.UI
{
    public class Menu : IDisposable
    {
        internal class AxHostConverter : AxHost
        {
            private AxHostConverter() : base("") { }

            static public IPictureDisp ImageToPictureDisp(Image image)
            {
                return (IPictureDisp)GetIPictureDispFromPicture(image);
            }

            static public IPicture ImageToPicture(Image image)
            {
                return (IPicture)GetIPictureFromPicture(image);
            }

            static public Image PictureDispToImage(IPictureDisp pictureDisp)
            {
                return GetPictureFromIPicture(pictureDisp);
            }
        }

        private readonly VBE _vbe;
        protected readonly AddIn AddIn;

        protected VBE IDE { get { return this._vbe; } }

        protected Menu(VBE vbe, AddIn addIn)
        {
            AddIn = addIn;
            _vbe = vbe;
        }

        private CommandBarButton AddButton(CommandBarPopup parentMenu, string caption)
        {
            var button = parentMenu.Controls.Add(MsoControlType.msoControlButton, null, null, null, true) as CommandBarButton;
            button.Caption = caption;

            return button;
        }

        protected CommandBarButton AddButton(CommandBarPopup parentMenu, string caption, bool beginGroup, CommandBarButtonClickEvent buttonClickHandler)
        {
            var button = AddButton(parentMenu, caption);
            button.BeginGroup = beginGroup;
            button.ClickEvent += buttonClickHandler;

            return button;
        }

        protected CommandBarButton AddButton(CommandBarPopup parentMenu, string caption, bool beginGroup, CommandBarButtonClickEvent buttonClickHandler, int faceId)
        {
            var button = AddButton(parentMenu, caption, beginGroup, buttonClickHandler);
            button.FaceId = faceId;

            return button;
        }

        protected CommandBarButton AddButton(CommandBarPopup parentMenu, string caption, bool beginGroup, CommandBarButtonClickEvent buttonClickHandler, Bitmap image)
        {
            var button = AddButton(parentMenu, caption, beginGroup, buttonClickHandler);
            SetButtonImage(button, image);

            return button;
        }

        public static void SetButtonImage(CommandBarButton button, Bitmap image)
        {
            button.FaceId = 0;

            if (image != null)
            {
                image.MakeTransparent(Color.Transparent);
                Clipboard.SetDataObject(image, true);
                button.PasteFace();
            }
        }

        public static void SetButtonImage(CommandBarButton button, Bitmap image, Bitmap mask)
        {
            button.FaceId = 0;
            button.Picture = (Picture)AxHostConverter.ImageToPicture(image);
            button.Mask = (Picture)AxHostConverter.ImageToPicture(mask);
        }

        /// <summary>
        /// Finds the index for insertion in a given CommandBarControls collection.
        /// Returns the last position if the given beforeControl caption is not found.
        /// </summary>
        /// <param name="controls">The collection to insert into.</param>
        /// <param name="beforeId">Caption of the control to insert before.</param>
        /// <returns></returns>
        protected int FindMenuInsertionIndex(CommandBarControls controls, int beforeId)
        {
            for (var i = 1; i <= controls.Count; i++)
            {
                if (controls[i].BuiltIn && controls[i].Id == beforeId)
                {
                    return i;
                }
            }

            return controls.Count;
        }

        public void Dispose()
        {
            Dispose(true);
        }

        protected virtual void Dispose(bool disposing)
        {
        }
    }
}
