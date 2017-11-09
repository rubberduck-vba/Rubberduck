using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
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
            var child = _items.FirstOrDefault(kvp => kvp.Value != null && kvp.Value.Tag == tag);
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

            foreach (var kvp in _items.Where(kv => kv.Key != null && kv.Value != null))
            {
                try
                {
                    var item = kvp;
                    UiDispatcher.Invoke(() =>
                    {
                        item.Value.Caption = item.Key.Caption.Invoke();
                        item.Value.TooltipText = item.Key.ToolTipText.Invoke();
                    });

                }
                catch (Exception e)
                {
                    Logger.Error(e, $"Assignment of {kvp.Value.GetType().Name}.Caption or .TooltipText for {kvp.Key.GetType().Name} threw an exception.");
                }
            }
        }

        public virtual void Initialize()
        {
            if (Parent == null)
            {
                return;
            }

            Item = Parent.Add(_name, _position);
            Item.IsVisible = true;
            foreach (var item in _items.Keys.OrderBy(item => item.DisplayOrder))
            {
                try
                {
                    _items[item] = InitializeChildControl(item);

                }
                catch (Exception e)
                {
                    Logger.Error(e, $"Initialization of the menu item for {item.Command.GetType().Name} threw an exception.");
                }
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
            foreach (var kvp in _items.Where(kv => kv.Key != null && kv.Value != null))
            {
                var commandItem = kvp.Key;
                var canExecute = false;
                try
                {
                    canExecute = commandItem.EvaluateCanExecute(state);

                }
                catch (Exception e)
                {
                    Logger.Error(e, $"{commandItem?.GetType().Name ?? nameof(ICommandMenuItem)}.EvaluateCanExecute(RubberduckParserState) threw an exception.");
                }
                kvp.Value.IsEnabled = canExecute;
                if (commandItem?.HiddenWhenDisabled ?? false)
                {
                    kvp.Value.IsVisible = canExecute;
                }
            }
        }

        public ICommandBars Parent { get; set; }
        public ICommandBar Item { get; private set; }

        public void RemoveCommandBar()
        {
            Logger.Debug("Removing commandbar.");
            RemoveChildren();
            Item?.Delete();
            Item = null;
            Parent = null;
        }

        private void RemoveChildren()
        {
            if (Parent == null || Parent.IsWrappingNullReference)
            {
                return;
            }

            try
            {
                foreach (var button in _items.Values.Select(item => item as ICommandBarButton).Where(child => child != null))
                {
                    if (!button.IsWrappingNullReference)
                    {
                        button.Click -= child_Click;
                    }
                    button.Delete();
                }
            }
            catch (COMException exception)
            {
                Logger.Error(exception, "Error removing child controls from commandbar.");
            }
            _items.Clear();
        }

        private void child_Click(object sender, CommandBarButtonClickEventArgs e)
        {
            var item = _items.Select(kvp => kvp.Key).SingleOrDefault(menu => menu.GetType().FullName == e.Control.Tag);
            if (item == null)
            {
                return;
            }

            Logger.Debug("({0}) Executing click handler for commandbar item '{1}', hash code {2}", GetHashCode(), e.Control.Caption, e.Control.Target.GetHashCode());
            item.Command.Execute(null);
        }
    }
}