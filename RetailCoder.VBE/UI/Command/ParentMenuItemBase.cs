using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Core;
using stdole;

namespace Rubberduck.UI.Command
{
    public abstract class ParentMenuItemBase : IParentMenuItem
    {
        private readonly CommandBarPopup _item;
        private readonly IDictionary<IMenuItem, CommandBarControl> _items;

        protected ParentMenuItemBase(CommandBarControls parent, string key, IEnumerable<IMenuItem> items, int? beforeIndex = null)
        {
            _item = beforeIndex.HasValue
                ? (CommandBarPopup) parent.Add(MsoControlType.msoControlPopup, Temporary: true, Before: beforeIndex)
                : (CommandBarPopup) parent.Add(MsoControlType.msoControlPopup, Temporary: true);
            _item.Tag = key;
            _items = items.ToDictionary(item => item, item => null as CommandBarControl);
        }

        public CommandBarPopup Item { get { return _item; } }

        public string Key { get { return _item.Tag; } }
        public Func<string> Caption { get { return () => RubberduckUI.ResourceManager.GetString(Key, RubberduckUI.Culture); } }

        public virtual bool BeginGroup { get { return false; } }
        public virtual int DisplayOrder { get { return default(int); } }
        
        public void Localize()
        {
            foreach (var kvp in _items)
            {
                kvp.Value.Caption = kvp.Key.Caption.Invoke();
            }
        }

        public void Initialize()
        {
            foreach (var item in _items.Keys.OrderBy(item => item.DisplayOrder))
            {
                _items[item] = InitializeChildControl(item as ICommandMenuItem)
                            ?? InitializeChildControl(item as IParentMenuItem);
            }
        }

        private CommandBarControl InitializeChildControl(IParentMenuItem item)
        {
            if (item == null)
            {
                return null;
            }

            item.Initialize();
            return item.Item;
        }

        private CommandBarControl InitializeChildControl(ICommandMenuItem item)
        {
            if (item == null)
            {
                return null;
            }

            var child = (CommandBarButton)_item.Controls.Add(MsoControlType.msoControlButton, Temporary: true);
            SetButtonImage(child, item.Image, item.Mask);

            child.Tag = item.Key;
            child.Caption = item.Caption.Invoke();
            child.Click += delegate { item.Command.Execute(); };

            return child;
        }

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
            private AxHostConverter() : base(string.Empty) { }

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
