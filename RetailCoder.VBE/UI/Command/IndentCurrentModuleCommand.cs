using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.SmartIndenter;

namespace Rubberduck.UI.Command
{
    [ComVisible(false)]
    public class IndentCurrentModuleCommand : CommandBase
    {
        private readonly VBE _vbe;
        private readonly RubberduckParserState _state;
        private readonly IIndenter _indenter;

        public IndentCurrentModuleCommand(VBE vbe, RubberduckParserState state, IIndenter indenter)
        {
            _vbe = vbe;
            _state = state;
            _indenter = indenter;
        }

        public override bool CanExecute(object parameter)
        {
            var target = FindTarget(parameter);

            return _vbe.ActiveCodePane != null && target != null &&
                   target.Annotations.All(a => a.AnnotationType != AnnotationType.NoIndent);
        }

        public override void Execute(object parameter)
        {
            _indenter.IndentCurrentModule();
        }

        public RubberduckHotkey Hotkey { get { return RubberduckHotkey.IndentModule; } }

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
