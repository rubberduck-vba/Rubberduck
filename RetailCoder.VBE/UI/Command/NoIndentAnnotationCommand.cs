using System.Linq;
using System.Runtime.InteropServices;
using NLog;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.Command
{
    [ComVisible(false)]
    public class NoIndentAnnotationCommand : CommandBase
    {
        private readonly IVBE _vbe;
        private readonly RubberduckParserState _state;

        public NoIndentAnnotationCommand(IVBE vbe, RubberduckParserState state) 
            : base(LogManager.GetCurrentClassLogger())
        {
            _vbe = vbe;
            _state = state;
        }

        protected override bool EvaluateCanExecute(object parameter)
        {
            var target = FindTarget(parameter);
            var pane = _vbe.ActiveCodePane;
            {
                return !pane.IsWrappingNullReference && target != null &&
                   target.Annotations.All(a => a.AnnotationType != AnnotationType.NoIndent);
            }
        }

        protected override void OnExecute(object parameter)
        {
            _vbe.ActiveCodePane.CodeModule.InsertLines(1, "'@NoIndent");
        }

        private Declaration FindTarget(object parameter)
        {
            if (parameter is Declaration declaration)
            {
                return declaration;
            }

            var selectedDeclaration = _state.FindSelectedDeclaration(_vbe.ActiveCodePane);

            while (selectedDeclaration != null && selectedDeclaration.DeclarationType.HasFlag(DeclarationType.Module))
            {
                selectedDeclaration = selectedDeclaration.ParentDeclaration;
            }

            return selectedDeclaration;
        }
    }
}
