using System;
using System.Drawing;

namespace Rubberduck.UI.Commands
{
    public interface ICommandMenuItem : IMenuItem
    {
        ICommand Command { get; }
    }

    public abstract class CommandMenuItemBase : ICommandMenuItem
    {
        private readonly ICommand _command;

        protected CommandMenuItemBase(ICommand command)
        {
            _command = command;
        }

        public abstract string Key { get; }

        public ICommand Command { get { return _command; } }
        public Func<string> Caption { get { return () => RubberduckUI.ResourceManager.GetString(Key, RubberduckUI.Culture); } }
        public bool IsParent { get { return false; } }
        public virtual Image Image { get { return null; } }
        public virtual Image Mask { get { return null; } }

    }
}