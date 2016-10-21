using System;
using System.Drawing;
using Castle.Core.Internal;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.UI.Command.MenuItems
{
    public abstract class CommandMenuItemBase : ICommandMenuItem
    {
        protected CommandMenuItemBase(CommandBase command)
        {
            _command = command;
        }

        private readonly CommandBase _command;
        public CommandBase Command { get { return _command; } }

        public abstract string Key { get; }

        public Func<string> Caption
        {
            get
            {
                return () => Key.IsNullOrEmpty() 
                    ? string.Empty 
                    : RubberduckUI.ResourceManager.GetString(Key, UI.Settings.Settings.Culture);
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
            return state != null && _command.CanExecute(state);
        }

        public virtual bool BeginGroup { get { return false; } }
        public virtual int DisplayOrder { get { return default(int); } }
        public virtual Image Image { get { return null; } }
        public virtual Image Mask { get { return null; } }
    }
}
