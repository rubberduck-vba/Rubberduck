using System;
using System.Drawing;
using System.Windows.Input;
using Castle.Core.Internal;
using Rubberduck.Parsing.VBA;

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

        /// <summary>
        /// Evaluates current parser state to determine whether command can be executed.
        /// </summary>
        /// <param name="state">The current parser state.</param>
        /// <returns>Returns <c>true</c> if command can be executed.</returns>
        /// <remarks>Returns <c>true</c> if not overridden.</remarks>
        public virtual bool EvaluateCanExecute(RubberduckParserState state)
        {
            return _command.CanExecute(state);
        }

        public virtual bool BeginGroup { get { return false; } }
        public virtual int DisplayOrder { get { return default(int); } }
        public virtual Image Image { get { return null; } }
        public virtual Image Mask { get { return null; } }
    }
}
