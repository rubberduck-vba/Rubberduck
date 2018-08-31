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
            using (var pane = _vbe.ActiveCodePane)
            {
                return !pane.IsWrappingNullReference && target != null &&
                       target.Annotations.All(a => a.AnnotationType != AnnotationType.NoIndent);
            }
        }

        protected override void OnExecute(object parameter)
        {
            using (var activePane = _vbe.ActiveCodePane)
            {
                if (activePane == null || activePane.IsWrappingNullReference)
                {
                    return;
                }

                using (var codeModule = activePane.CodeModule)
                {
                    codeModule.InsertLines(1, "'@NoIndent");
                }
            }
        }

        private Declaration FindTarget(object parameter)
        {
            if (parameter is Declaration declaration)
            {
                return declaration;
            }

            Declaration selectedDeclaration;
            using (var activePane = _vbe.ActiveCodePane)
            {
                selectedDeclaration = _state.FindSelectedDeclaration(activePane);
            }

            while (selectedDeclaration != null && selectedDeclaration.DeclarationType.HasFlag(DeclarationType.Module))
            {
                selectedDeclaration = selectedDeclaration.ParentDeclaration;
            }

            return selectedDeclaration;
        }
    }
}
