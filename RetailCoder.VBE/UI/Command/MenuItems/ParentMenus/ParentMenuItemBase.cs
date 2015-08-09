using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Core;
using stdole;

namespace Rubberduck.UI.Command.MenuItems.ParentMenus
{
    public abstract class ParentMenuItemBase : IParentMenuItem
    {
        private readonly string _key;
        private readonly int? _beforeIndex;
        private readonly IDictionary<IMenuItem, CommandBarControl> _items;

        protected ParentMenuItemBase(string key, IEnumerable<IMenuItem> items, int? beforeIndex = null)
        {
            _key = key;
            _beforeIndex = beforeIndex;
            _items = items.ToDictionary(item => item, item => null as CommandBarControl);
        }

        public CommandBarControls Parent { get; set; }
        public CommandBarPopup Item { get; private set; }

        public string Key { get { return Item == null ? null : Item.Tag; } }

        public Func<string> Caption { get { return () => Key == null ? null : RubberduckUI.ResourceManager.GetString(Key, RubberduckUI.Culture); } }

        public virtual bool BeginGroup { get { return false; } }
        public virtual int DisplayOrder { get { return default(int); } }
        
        public void Localize()
        {
            if (Item == null)
            {
                return;
            }

            Item.Caption = Caption.Invoke();
            foreach (var kvp in _items)
            {
                kvp.Value.Caption = kvp.Key.Caption.Invoke();
            }
        }

        public void Initialize()
        {
            if (Parent == null)
            {
                return;
            }

            Item = _beforeIndex.HasValue
                ? (CommandBarPopup)Parent.Add(MsoControlType.msoControlPopup, Temporary: true, Before: _beforeIndex)
                : (CommandBarPopup)Parent.Add(MsoControlType.msoControlPopup, Temporary: true);
            Item.Tag = _key;

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

            item.Parent = Item.Controls;
            item.Initialize();
            return item.Item;
        }

        private CommandBarControl InitializeChildControl(ICommandMenuItem item)
        {
            if (item == null)
            {
                return null;
            }

            var child = (CommandBarButton)Item.Controls.Add(MsoControlType.msoControlButton, Temporary: true);
            SetButtonImage(child, item.Image, item.Mask);

            child.BeginGroup = item.BeginGroup;
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
