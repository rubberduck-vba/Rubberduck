using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UnitTesting;

namespace Rubberduck.UI.Command
{
    /// <summary>
    /// A command that adds a new test method stub to the active code pane.
    /// </summary>
    [ComVisible(false)]
    public class AddTestMethodExpectedErrorCommand : CommandBase
    {
        private readonly VBE _vbe;
        private readonly NewTestMethodCommand _command;
        private readonly RubberduckParserState _state;

        public AddTestMethodExpectedErrorCommand(VBE vbe, RubberduckParserState state, NewTestMethodCommand command)
        {
            _vbe = vbe;
            _command = command;
            _state = state;
        }

        public override bool CanExecute(object parameter)
        {
            if (_state.Status != ParserState.Ready) { return false; }

            var testModules = _state.AllUserDeclarations.Where(d =>
                        d.DeclarationType == DeclarationType.ProceduralModule &&
                        d.Annotations.Any(a => a.AnnotationType == AnnotationType.TestModule));

            return testModules.Any(a => a.QualifiedName.QualifiedModuleName.Component == _vbe.SelectedVBComponent);
        }

        public override void Execute(object parameter)
        {
            _command.NewExpectedErrorTestMethod();
        }
    }
}
