using System.Runtime.InteropServices;
using NLog;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.Command
{
    [ComVisible(false)]
    public class IndentCurrentProcedureCommand : CommandBase
    {
        private readonly IVBE _vbe;
        private readonly IIndenter _indenter;
        private readonly RubberduckParserState _state;

        public IndentCurrentProcedureCommand(IVBE vbe, IIndenter indenter, RubberduckParserState state)
            : base(LogManager.GetCurrentClassLogger())
        {
            _vbe = vbe;
            _indenter = indenter;
            _state = state;
        }

        public override RubberduckHotkey Hotkey
        {
            get { return RubberduckHotkey.IndentProcedure; }
        }

        protected override bool CanExecuteImpl(object parameter)
        {
            return _vbe.ActiveCodePane != null;
        }

        protected override void ExecuteImpl(object parameter)
        {
            _indenter.IndentCurrentProcedure();
            _state.OnParseRequested(this, _vbe.ActiveCodePane.CodeModule.Parent);
        }
    }
}
