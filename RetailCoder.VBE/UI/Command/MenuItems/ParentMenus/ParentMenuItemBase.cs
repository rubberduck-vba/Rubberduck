using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Rubberduck.Parsing.VBA;
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

            Debug.Print("'{0}' ({1}) parent menu initialized, hash code {2}.", _key, GetHashCode(), Item.GetHashCode());
        }

        public void EvaluateCanExecute(RubberduckParserState state)
        {
            foreach (var kvp in _items)
            {
                var parentItem = kvp.Key as IParentMenuItem;
                if (parentItem != null)
                {
                    parentItem.EvaluateCanExecute(state);
                    continue;
                }

                var commandItem = kvp.Key as ICommandMenuItem;
                if (commandItem != null)
                {
                    kvp.Value.Enabled = commandItem.EvaluateCanExecute(state);
                }
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

            Debug.WriteLine("Menu item '{0}' created; hash code: {1} (command hash code {2})", child.Caption, child.GetHashCode(), item.Command.GetHashCode());

            child.Click += child_Click;
            return child;
        }

        // note: HAAAAACK!!!
        private static int _lastHashCode;

        private void child_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            var item = _items.Select(kvp => kvp.Key).SingleOrDefault(menu => menu.Key == Ctrl.Tag) as ICommandMenuItem;
            if (item == null || Ctrl.GetHashCode() == _lastHashCode)
            {
                return;
            }

            // without this hack, handler runs once for each menu item that's hooked up to the command.
            // hash code is different on every frakkin' click. go figure. I've had it, this is the fix.
            _lastHashCode = Ctrl.GetHashCode();

            Debug.WriteLine("({0}) Executing click handler for menu item '{1}', hash code {2}", GetHashCode(), Ctrl.Caption, Ctrl.GetHashCode());
            item.Command.Execute(null);
        }

        /// <summary>
        /// Creates a transparent <see cref="IPictureDisp"/> icon for the specified <see cref="CommandBarButton"/>.
        /// </summary>
        public static void SetButtonImage(CommandBarButton button, Image image, Image mask)
        {
            button.FaceId = 0;
            if (image == null || mask == null)
            {
                return;
            }

            try
            {
                button.Picture = AxHostConverter.ImageToPictureDisp(image);
                button.Mask = AxHostConverter.ImageToPictureDisp(mask);
            }
            catch (COMException exception)
            {
                Debug.Print("Button image could not be set for button [" + button.Caption + "]\n" + exception);
            }
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
