using System.Collections.Generic;
using System.Linq;
using NLog;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.MenuItems.ParentMenus;
using Rubberduck.VBEditor.SafeComWrappers.Office.Core;
using Rubberduck.VBEditor.SafeComWrappers.Office.Core.Abstract;

namespace Rubberduck.UI.Command.MenuItems.CommandBars
{
    public abstract class AppCommandBarBase : IAppCommandBar
    {
        private readonly string _name;
        private readonly CommandBarPosition _position;
        private readonly IDictionary<ICommandMenuItem, ICommandBarControl> _items;
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        protected AppCommandBarBase(string name, CommandBarPosition position, IEnumerable<ICommandMenuItem> items)
        {
            _name = name;
            _position = position;
            _items = items.ToDictionary(item => item, item => null as ICommandBarControl);
        }

        protected ICommandMenuItem FindChildByTag(string tag)
        {
            var child = _items.FirstOrDefault(kvp => kvp.Value.Tag == tag);
            return Equals(child, default(KeyValuePair<ICommandMenuItem, ICommandBarControl>)) 
                ? null 
                : child.Key;
        }

        public void Localize()
        {
            if (Item == null)
            {
                return;
            }

            foreach (var kvp in _items)
            {
                var item = kvp;
                UiDispatcher.Invoke(() =>
                {
                    item.Value.Caption = item.Key.Caption.Invoke();
                    item.Value.TooltipText = item.Key.ToolTipText.Invoke();
                });
            }
        }

        public void Initialize()
        {
            if (Parent == null)
            {
                return;
            }

            Item = Parent.Add(_name, _position);
            Item.IsVisible = true;
            foreach (var item in _items.Keys.OrderBy(item => item.DisplayOrder))
            {
                _items[item] = InitializeChildControl(item);
            }
        }

        private ICommandBarControl InitializeChildControl(ICommandMenuItem item)
        {
            if (item == null)
            {
                return null;
            }

            var child = CommandBarButtonFactory.Create(Item.Controls);
            child.Style = item.ButtonStyle;
            child.Picture = item.Image;
            child.Mask = item.Mask;
            child.ApplyIcon();

            child.IsVisible = item.IsVisible;
            child.BeginsGroup = item.BeginGroup;
            child.Tag = item.GetType().FullName;
            child.Caption = item.Caption.Invoke();
            child.TooltipText = item.ToolTipText.Invoke();

            if (item.Command != null)
            {
                child.Click += child_Click;
            }
            return child;
        }

        public void EvaluateCanExecute(RubberduckParserState state)
        {
            foreach (var kvp in _items)
            {
                var commandItem = kvp.Key;
                if (commandItem != null && kvp.Value != null)
                {
                    var canExecute = commandItem.EvaluateCanExecute(state);
                    kvp.Value.IsEnabled = canExecute;
                    if (commandItem.HiddenWhenDisabled)
                    {
                        kvp.Value.IsVisible = canExecute;
                    }
                }
            }
        }

        public ICommandBars Parent { get; set; }
        public ICommandBar Item { get; private set; }
        public void RemoveChildren()
        {
            // note: doing this wrecks the teardown process. counter-intuitive? sure. but hey it works.
            foreach (var child in _items.Values.Select(item => item as CommandBarButton).Where(child => child != null))
            {
                child.Click -= child_Click;
            //    child.Delete();
            //    child.Release();
            }
        }

        // note: HAAAAACK!!!
        private static int _lastHashCode;

        private void child_Click(object sender, CommandBarButtonClickEventArgs e)
        {
            var item = _items.Select(kvp => kvp.Key).SingleOrDefault(menu => menu.GetType().FullName == e.Control.Tag);
            if (item == null || e.Control.Target.GetHashCode() == _lastHashCode)
            {
                return;
            }

            // without this hack, handler runs once for each menu item that's hooked up to the command.
            // hash code is different on every frakkin' click. go figure. I've had it, this is the fix.
            _lastHashCode = e.Control.Target.GetHashCode();

            Logger.Debug("({0}) Executing click handler for commandbar item '{1}', hash code {2}", GetHashCode(), e.Control.Caption, e.Control.Target.GetHashCode());
            item.Command.Execute(null);
        }
    }
}