using System;
using System.Drawing;
using Castle.Core.Internal;

namespace Rubberduck.UI.Command.MenuItems
{
    public abstract class CommandMenuItemBase : ICommandMenuItem
    {
        protected CommandMenuItemBase(ICommand command)
        {
            _command = command;
        }

        private readonly ICommand _command;
        public ICommand Command { get { return _command; } }

        public abstract string Key { get; }

        public Func<string> Caption
        {
            get
            {
                return () => Key.IsNullOrEmpty() 
                    ? string.Empty 
                    : RubberduckUI.ResourceManager.GetString(Key, RubberduckUI.Culture);
            }
        }

        public virtual bool BeginGroup { get { return false; } }
        public virtual int DisplayOrder { get { return default(int); } }
        public virtual Image Image { get { return null; } }
        public virtual Image Mask { get { return null; } }
    }
}
