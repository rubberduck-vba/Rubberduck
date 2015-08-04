using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Core;
using stdole;

namespace Rubberduck.UI.Commands
{
    /// <summary>
    /// An objects that wraps a <see cref="CommandBarPopup"/> instance.
    /// </summary>
    public abstract class ParentMenu : IMenuItem, ICommandBar
    {
        private readonly string _key;
        private readonly Func<string> _caption;
        private readonly CommandBarPopup _popup;
        private readonly IDictionary<IMenuItem, CommandBarControl> _items = new Dictionary<IMenuItem, CommandBarControl>();

        protected ParentMenu(CommandBarControls parent, string key, Func<string> caption, int? beforeIndex)
        {
            _key = key;
            _caption = caption;
            _popup = beforeIndex.HasValue
                ? (CommandBarPopup) parent.Add(MsoControlType.msoControlPopup, Temporary: true, Before: beforeIndex)
                : (CommandBarPopup) parent.Add(MsoControlType.msoControlPopup, Temporary: true);
            _popup.Tag = _key;

            Localize();
        }

        public abstract void Initialize();

        public string Key { get { return _key; } }
        public Func<string> Caption { get {return _caption; } }
        public bool IsParent { get { return true; } }
        public Image Image { get {return null; } }
        public Image Mask { get { return null; } }

        public void Localize()
        {
            _popup.Caption = _caption.Invoke();
            LocalizeChildren();
        }

        private void LocalizeChildren()
        {
            foreach (var kvp in _items)
            {
                var value = kvp.Key.Caption.Invoke();
                kvp.Value.Caption = value;
            }
        }

        public void AddItem(IMenuItem item, bool? beginGroup = null, int? beforeIndex = null)
        {
            var controlType = item.IsParent 
                ? MsoControlType.msoControlPopup 
                : MsoControlType.msoControlButton;
            var child = beforeIndex.HasValue
                ? _popup.Controls.Add(controlType, Temporary: true, Before: beforeIndex)
                : _popup.Controls.Add(controlType, Temporary: true);

            child.Caption = item.Caption.Invoke();
            child.BeginGroup = beginGroup ?? false;
            child.Tag = item.Key;

            if (!item.IsParent)
            {
                var button = (CommandBarButton)child;
                SetButtonImage(button, item.Image, item.Mask);

                var command = ((ICommandMenuItem)item).Command;
                button.Click += delegate { command.Execute(); };
            }

            _items.Add(item, child);
        }

        public bool RemoveItem(IMenuItem item)
        {
            try
            {
                var child = _items[item];
                child.Delete();
                Marshal.ReleaseComObject(child);
                _items.Remove(item);
                return true;
            }
            catch (COMException)
            {
                return false;
            }
        }

        public bool Remove()
        {
            foreach (var menuItem in _items)
            {
                RemoveItem(menuItem.Key); // note: should we care if this fails?
            }

            try
            {
                _popup.Delete();
                Marshal.ReleaseComObject(_popup);
                return true;
            }
            catch (COMException)
            {
                return false;
            }
        }

        public IEnumerable<IMenuItem> Items { get { return _items.Keys; } }

        private static void SetButtonImage(CommandBarButton button, Image image, Image mask)
        {
            button.FaceId = 0;
            if (image == null || mask == null)
            {
                return;
            }

            button.Picture = AxHostConverter.ImageToPictureDisp(image);
            button.Mask = AxHostConverter.ImageToPictureDisp(mask);
        }

        private class AxHostConverter : AxHost
        {
            private AxHostConverter() : base("") { }

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
}