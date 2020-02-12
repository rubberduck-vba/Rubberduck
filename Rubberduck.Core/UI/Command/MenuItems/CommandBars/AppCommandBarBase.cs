using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using NLog;
using Rubberduck.Parsing.UIContext;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.Command.MenuItems.CommandBars
{
    public abstract class AppCommandBarBase : IAppCommandBar
    {
        private readonly string _name;
        private readonly CommandBarPosition _position;
        private readonly IDictionary<ICommandMenuItem, ICommandBarControl> _items;
        protected readonly IUiDispatcher _uiDispatcher;
        protected static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        protected AppCommandBarBase(string name, CommandBarPosition position, IEnumerable<ICommandMenuItem> items, IUiDispatcher uiDispatcher)
        {
            _name = name;
            _position = position;
            _items = items.ToDictionary(item => item, item => null as ICommandBarControl);
            _uiDispatcher = uiDispatcher;
        }

        protected ICommandMenuItem FindChildByTag(string tag)
        {
            try
            {
                var child = _items.FirstOrDefault(kvp => kvp.Value != null && kvp.Value?.Tag == tag);
                return Equals(child, default(KeyValuePair<ICommandMenuItem, ICommandBarControl>))
                    ? null
                    : child.Key;
            }
            catch (COMException exception)
            {
                Logger.Error(exception, $"COMException while finding child with tag '{tag}' in the command bar.");
            }
            catch (InvalidCastException exception)
            {
                //This exception will be encountered whenever the registration of the command bar control not correct in the system.
                //See issue #5349 at https://github.com/rubberduck-vba/Rubberduck/issues/5349
                Logger.Error(exception, $"Invalid cast exception while finding child with tag '{tag}' in the command bar.");
            }
            return null;
        }

        public void Localize()
        {
            if (Item == null)
            {
                return;
            }

            foreach (var kvp in _items.Where(kv => kv.Key != null && kv.Value != null && !kv.Value.IsWrappingNullReference))
            {
                try
                {
                    var item = kvp;
                    _uiDispatcher.Invoke(() => LocalizeInternal(item, kvp));

                }
                catch (Exception e)
                {
                    Logger.Error(e, $"Failed to dispatch assignment of {kvp.Value.GetType().Name}.Caption and .TooltipText for {kvp.Key.GetType().Name} to the UI thread.");
                }
            }
        }

        private static void LocalizeInternal(KeyValuePair<ICommandMenuItem, ICommandBarControl> item, KeyValuePair<ICommandMenuItem, ICommandBarControl> kvp)
        {
            try
            {
                item.Value.Caption = item.Key.Caption.Invoke();
                item.Value.TooltipText = item.Key.ToolTipText.Invoke();
            }
            catch (Exception e)
            {
                Logger.Error(e,
                    $"Assignment of {kvp.Value.GetType().Name}.Caption or .TooltipText for {kvp.Key.GetType().Name} threw an exception.");
            }
        }

        public virtual void Initialize()
        {
            if (Parent == null  || Parent.IsWrappingNullReference)
            {
                return;
            }

            try
            {
                Item = Parent.Add(_name, _position);
                Item.IsVisible = true;
            }
            catch (COMException exception)
            {
                Logger.Error(exception, $"Failed to add the command bar {_name}.");
                return;
            }
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

            ICommandBarButton child;
            using (var controls = Item.Controls)
            {
                child = controls.AddButton();
            }
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
            foreach (var kvp in _items.Where(kv => kv.Key != null && kv.Value != null && !kv.Value.IsWrappingNullReference))
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
                try
                {
                    kvp.Value.IsEnabled = canExecute;
                    if (commandItem?.HiddenWhenDisabled ?? false)
                    {
                        kvp.Value.IsVisible = canExecute;
                    }
                }
                catch (COMException exception)
                {
                    Logger.Error(exception,$"COMException while trying to set IsEnabled and IsVisible on {commandItem?.GetType().Name ?? nameof(ICommandMenuItem)}");
                }
            }
        }

        private ICommandBars _parent;
        public ICommandBars Parent
        {
            get => _parent;
            set
            {
                _parent?.Dispose();
                _parent = value;
            }
        }

        private ICommandBar _item;
        public ICommandBar Item
        {
            get => _item;
            private set
            {
                _item?.Dispose();
                _item = value;
            }
        }

        public void RemoveCommandBar()
        {
            try
            {
                if (Item != null)
                {
                    Logger.Debug("Removing commandbar.");
                    RemoveChildren();
                    Item.Delete();

                    // Setting them to null will automatically dispose those
                    Item = null;
                    Parent = null;
                }
            }
            catch (COMException exception)
            {
                Logger.Error(exception, "COM exception while trying to delete the commandbar");
            }
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
                    button.Dispose();
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
            ICommandMenuItem item;
            try
            {
                item = _items.Select(kvp => kvp.Key).SingleOrDefault(menu => menu.GetType().FullName == e.Tag);
            }
            catch (COMException exception)
            {
                Logger.Error(exception, "COM exception finding command for a control.");
                item = null;
            }
            if (item == null)
            {
                return;
            }

            Logger.Debug("({0}) Executing click handler for commandbar item '{1}', hash code {2}", GetHashCode(), e.Caption, e.TargetHashCode);
            item.Command.Execute(null);
        }
    }
}