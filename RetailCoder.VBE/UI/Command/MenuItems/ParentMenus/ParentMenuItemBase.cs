using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Rubberduck.Parsing.VBA;
using NLog;
using Rubberduck.VBEditor.SafeComWrappers.Office.Core;

namespace Rubberduck.UI.Command.MenuItems.ParentMenus
{
    public abstract class ParentMenuItemBase : IParentMenuItem
    {
        private readonly string _key;
        private readonly int? _beforeIndex;
        private readonly IDictionary<IMenuItem, CommandBarControl> _items;
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        protected ParentMenuItemBase(string key, IEnumerable<IMenuItem> items, int? beforeIndex = null)
        {
            _key = key;
            _beforeIndex = beforeIndex;
            _items = items.ToDictionary(item => item, item => null as CommandBarControl);
        }

        public CommandBarControls Parent { get; set; }
        public CommandBarPopup Item { get; private set; }

        public string Key { get { return Item == null ? null : Item.Tag; } }

        public Func<string> Caption { get { return () => Key == null ? null : RubberduckUI.ResourceManager.GetString(Key, Settings.Settings.Culture); } }

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
                var command = kvp.Key as CommandMenuItemBase;
                if (command != null)
                {
                    ((CommandBarButton)kvp.Value).ShortcutText = command.Command.ShortcutText;
                }

                var childMenu = kvp.Key as ParentMenuItemBase;
                if (childMenu != null)
                {
                    childMenu.Localize();
                }
            }
        }

        public void Initialize()
        {
            if (Parent == null)
            {
                return;
            }

            Item = _beforeIndex.HasValue
                ? CommandBarPopup.FromCommandBarControl(Parent.Add(ControlType.Popup, _beforeIndex.Value))
                : CommandBarPopup.FromCommandBarControl(Parent.Add(ControlType.Popup));
            Item.Tag = _key;

            foreach (var item in _items.Keys.OrderBy(item => item.DisplayOrder))
            {
                _items[item] = InitializeChildControl(item as ICommandMenuItem)
                            ?? InitializeChildControl(item as IParentMenuItem);
            }
        }

        public void RemoveChildren()
        {
            foreach (var child in _items.Keys.Select(item => item as IParentMenuItem).Where(child => child != null))
            {
                child.RemoveChildren();
                var item = _items[child];
                Debug.Assert(item is CommandBarPopup);
                (item as CommandBarPopup).Delete();
            }
            foreach (var child in _items.Values.Select(item => item as CommandBarButton).Where(child => child != null))
            {
                child.Click -= child_Click;
                child.Delete();
            }
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
                if (commandItem != null && kvp.Value != null)
                {
                     kvp.Value.IsEnabled = commandItem.EvaluateCanExecute(state);
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

            var child = CommandBarButton.FromCommandBarControl(Item.Controls.Add(ControlType.Button));
            child.Picture = item.Image;
            child.Mask = item.Mask;
            child.ApplyIcon();

            child.BeginsGroup = item.BeginGroup;
            child.Tag = item.GetType().FullName;
            child.Caption = item.Caption.Invoke();
            var command = item.Command as CommandBase; // todo: add 'ShortcutText' to a new 'interface CommandBase : System.Windows.Input.CommandBase'
            child.ShortcutText = command != null
                ? command.ShortcutText
                : string.Empty;

            child.Click += child_Click;
            return child;
        }

        // note: HAAAAACK!!!
        private static int _lastHashCode;

        private void child_Click(object sender, CommandBarButtonClickEventArgs e)
        {
            var item = _items.Select(kvp => kvp.Key).SingleOrDefault(menu => menu.GetType().FullName == e.Control.Tag) as ICommandMenuItem;
            if (item == null || e.Control.GetHashCode() == _lastHashCode)
            {
                return;
            }

            // without this hack, handler runs once for each menu item that's hooked up to the command.
            // hash code is different on every frakkin' click. go figure. I've had it, this is the fix.
            _lastHashCode = e.Control.GetHashCode();

            Logger.Debug("({0}) Executing click handler for menu item '{1}', hash code {2}", GetHashCode(), e.Control.Caption, e.Control.GetHashCode());
            item.Command.Execute(null);
        }
    }
}
