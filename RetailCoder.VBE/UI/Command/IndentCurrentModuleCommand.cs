using System.Runtime.InteropServices;
using NLog;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.Command
{
    [ComVisible(false)]
    public class IndentCurrentModuleCommand : CommandBase
    {
        private readonly IVBE _vbe;
        private readonly IIndenter _indenter;
        private readonly RubberduckParserState _state;

        public IndentCurrentModuleCommand(IVBE vbe, IIndenter indenter, RubberduckParserState state) : base(LogManager.GetCurrentClassLogger())
        {
            _vbe = vbe;
            _indenter = indenter;
            _state = state;
        }

        public override RubberduckHotkey Hotkey
        {
            get { return RubberduckHotkey.IndentModule; }
        }

        protected override bool CanExecuteImpl(object parameter)
        {
            return !_vbe.ActiveCodePane.IsWrappingNullReference;
        }

        protected override void ExecuteImpl(object parameter)
        {
            _indenter.IndentCurrentModule();
            _state.OnParseRequested(this, _vbe.ActiveCodePane.CodeModule.Parent);
        }
    }
}
