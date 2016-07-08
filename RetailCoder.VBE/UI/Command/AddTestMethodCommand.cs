using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using NLog;
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
    public class AddTestMethodCommand : CommandBase
    {
        private readonly VBE _vbe;
        private readonly NewTestMethodCommand _command;
        private readonly RubberduckParserState _state;

        public AddTestMethodCommand(VBE vbe, RubberduckParserState state, NewTestMethodCommand command) : base(LogManager.GetCurrentClassLogger())
        {
            _vbe = vbe;
            _command = command;
            _state = state;
        }

        protected override bool CanExecuteImpl(object parameter)
        {
            if (_state.Status != ParserState.Ready) { return false; }

            var testModules = _state.AllUserDeclarations.Where(d =>
                        d.DeclarationType == DeclarationType.ProceduralModule &&
                        d.Annotations.Any(a => a.AnnotationType == AnnotationType.TestModule));

            try
            {
                // the code modules consistently match correctly, but the components don't
                return testModules.Any(a =>
                            a.QualifiedName.QualifiedModuleName.Component.CodeModule ==
                            _vbe.SelectedVBComponent.CodeModule);
            }
            catch (COMException)
            {
                return false;
            }
        }

        protected override void ExecuteImpl(object parameter)
        {
            _command.NewTestMethod();
        }
    }
}
