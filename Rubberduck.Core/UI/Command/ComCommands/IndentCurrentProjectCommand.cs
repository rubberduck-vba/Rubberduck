using System.Runtime.InteropServices;
using Rubberduck.Parsing.VBA;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.Command.ComCommands
{
    [ComVisible(false)]
    public class IndentCurrentProjectCommand : ComCommandBase
    {
        private readonly IVBE _vbe;
        private readonly IIndenter _indenter;
        private readonly RubberduckParserState _state;

        public IndentCurrentProjectCommand(
            IVBE vbe, 
            IIndenter indenter, 
            RubberduckParserState state, 
            IVbeEvents vbeEvents) 
            : base(vbeEvents)
        {
            _vbe = vbe;
            _indenter = indenter;
            _state = state;

            AddToCanExecuteEvaluation(SpecialEvaluateCanExecute);
        }

        private bool SpecialEvaluateCanExecute(object parameter)
        {
            using (var vbProject = _vbe.ActiveVBProject)
            {
                return !vbProject.IsWrappingNullReference &&
                       vbProject.Protection != ProjectProtection.Locked;
            }
        }

        protected override void OnExecute(object parameter)
        {
            _indenter.IndentCurrentProject();
            if (_state.Status >= ParserState.Ready || _state.Status == ParserState.Pending)
            {
                _state.OnParseRequested(this);
            }
        }
    }
}
