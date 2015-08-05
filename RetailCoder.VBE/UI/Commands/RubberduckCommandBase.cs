using System;

namespace Rubberduck.UI.Commands
{
    /// <summary>
    /// Base class to derived all parameterized menu commands from.
    /// </summary>
    /// <typeparam name="TParam">The type of the parameter used by the command.</typeparam>
    public abstract class RubberduckParamCommandBase<TParam> : RubberduckCommandBase
    {
        protected RubberduckParamCommandBase(IRubberduckMenuCommand command)
            : base(command)
        {
        }

        /// <summary>
        /// A method that enables executing the command with a strongly-typed parameter.
        /// Base method simply calls non-parameterized <see cref="ExecuteAction"/> method.
        /// </summary>
        /// <param name="parameter">An object containing information needed to execute a parameterized command.</param>
        public virtual void ExecuteAction(TParam parameter)
        {
            ExecuteAction();
        }
    }

    /// <summary>
    /// Base class to derive all menu commands from.
    /// </summary>
    public abstract class RubberduckCommandBase
    {
        private readonly IRubberduckMenuCommand _command;

        protected RubberduckCommandBase(IRubberduckMenuCommand command)
        {
            _command = command;
            _command.RequestExecute += command_RequestExecute;
        }

        private void command_RequestExecute(object sender, EventArgs e)
        {
            ExecuteAction();
        }

        protected IRubberduckMenuCommand Command { get { return _command; } }

        /// <summary>
        /// A method that uses the <see cref="Command"/> helper to wire up as many UI controls as needed.
        /// </summary>
        public abstract void Initialize();

        /// <summary>
        /// The method that is executed when either wired-up UI control is clicked.
        /// </summary>
        public abstract void ExecuteAction();

        public void Release()
        {
            _command.Release();
        }
    }
}