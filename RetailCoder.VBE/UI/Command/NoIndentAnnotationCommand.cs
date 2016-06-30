using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using NLog;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.UI.Command
{
    [ComVisible(false)]
    public class NoIndentAnnotationCommand : CommandBase
    {
        private readonly VBE _vbe;
        private readonly RubberduckParserState _state;

        public NoIndentAnnotationCommand(VBE vbe, RubberduckParserState state) : base(LogManager.GetCurrentClassLogger())
        {
            _vbe = vbe;
            _state = state;
        }

        protected override bool CanExecuteImpl(object parameter)
        {
            var target = FindTarget(parameter);

            return _vbe.ActiveCodePane != null && target != null &&
                   target.Annotations.All(a => a.AnnotationType != AnnotationType.NoIndent);
        }

        protected override void ExecuteImpl(object parameter)
        {
            _vbe.ActiveCodePane.CodeModule.InsertLines(1, "'@NoIndent");
        }

        private Declaration FindTarget(object parameter)
        {
            var declaration = parameter as Declaration;
            if (declaration != null)
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
