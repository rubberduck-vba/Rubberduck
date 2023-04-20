using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.Interaction;
using Rubberduck.Parsing.Annotations.Concrete;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.Command.ComCommands
{
    [ComVisible(false)]
    public class IndentCurrentModuleCommand : ComCommandBase
    {
        private readonly IVBE _vbe;
        private readonly IIndenter _indenter;
        private readonly RubberduckParserState _state;
        private readonly IMessageBox _messageBox;

        public IndentCurrentModuleCommand(
            IVBE vbe, 
            IIndenter indenter, 
            RubberduckParserState state,
            IVbeEvents vbeEvents,
            IMessageBox messageBox) 
            : base(vbeEvents)
        {
            _vbe = vbe;
            _indenter = indenter;
            _state = state;
            _messageBox = messageBox;

            AddToCanExecuteEvaluation(SpecialEvaluateCanExecute);
        }

        private bool SpecialEvaluateCanExecute(object parameter)
        {
            using (var activePane = _vbe.ActiveCodePane)
            {
                return activePane != null && !activePane.IsWrappingNullReference;
            }
        }

        protected override void OnExecute(object parameter)
        {            
            if (_state.IsDirty())
            {
                if (!_messageBox.ConfirmYesNo(
                        Resources.RubberduckUI.Indenter_ContinueIndentWithoutAnnotations,
                        Resources.RubberduckUI.Indenter_ContinueIndentWithoutAnnotations_DialogCaption,
                        false))
                    return;

                _indenter.IndentCurrentModule();
            }
            else
            {                                
                var componentDeclarations = _state.AllUserDeclarations.Where(c =>
                    c.DeclarationType.HasFlag(DeclarationType.Module) &&
                    !c.Annotations.Any(pta => pta.Annotation is NoIndentAnnotation) &&
                    c.ProjectId == _vbe.ActiveVBProject.ProjectId &&
                    c.QualifiedModuleName == _vbe.ActiveCodePane.QualifiedModuleName
                    );

                foreach (var componentDeclaration in componentDeclarations)
                {
                    _indenter.Indent(_state.ProjectsProvider.Component(componentDeclaration.QualifiedName.QualifiedModuleName));
                }
            }

            if (_state.Status >= ParserState.Ready || _state.Status == ParserState.Pending)
            {
                _state.OnParseRequested(this);
            }
        }
    }
}
