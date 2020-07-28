using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.Parsing.Annotations.Concrete;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.ComCommands;
using Rubberduck.UnitTesting;
using Rubberduck.UnitTesting.CodeGeneration;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.UnitTesting.ComCommands
{
    /// <summary>
    /// A command that adds a new test method stub to the active code pane.
    /// </summary>
    [ComVisible(false)]
    public class AddTestMethodExpectedErrorCommand : ComCommandBase
    {
        private readonly IVBE _vbe;
        private readonly RubberduckParserState _state;
        private readonly ITestCodeGenerator _codeGenerator;

        public AddTestMethodExpectedErrorCommand(
            IVBE vbe, 
            RubberduckParserState state, 
            ITestCodeGenerator codeGenerator, 
            IVbeEvents vbeEvents) 
            : base(vbeEvents)
        {
            _vbe = vbe;
            _state = state;
            _codeGenerator = codeGenerator;

            AddToCanExecuteEvaluation(SpecialEvaluateCanExecute);
        }

        private bool SpecialEvaluateCanExecute(object parameter)
        {
            using (var pane = _vbe.ActiveCodePane)
            {
                if (_state.Status != ParserState.Ready || pane is null || pane.IsWrappingNullReference)
                {
                    return false;
                }
            }
            var testModules = _state.AllUserDeclarations.Where(d =>
                            d.DeclarationType == DeclarationType.ProceduralModule &&
                            d.Annotations.Any(pta => pta.Annotation is TestModuleAnnotation));

            try
            {
                // the code modules consistently match correctly, but the components don't
                using (var component = _vbe.SelectedVBComponent)
                {
                    using(var selectedModule = component.CodeModule)
                    {
                        return testModules.Any(a => _state.ProjectsProvider.Component(a.QualifiedModuleName).HasEqualCodeModule(selectedModule));
                    }
                }
            }
            catch (COMException)
            {
                return false;
            }      
        }

        protected override void OnExecute(object parameter)
        {
            using (var pane = _vbe.ActiveCodePane)
            {
                if (pane?.IsWrappingNullReference ?? true)
                {
                    return;
                }

                using (var activeModule = pane.CodeModule)
                {
                    var declaration = _state.GetTestModules().FirstOrDefault(f =>
                    {
                        var component = _state.ProjectsProvider.Component(f.QualifiedName.QualifiedModuleName);
                        using (var thisModule = component.CodeModule)
                        {
                            return thisModule.Equals(activeModule);
                        }
                    });

                    if (declaration != null)
                    {
                        using (var component = activeModule.Parent)
                        {
                            activeModule.InsertLines(activeModule.CountOfLines, _codeGenerator.GetNewTestMethodCodeErrorExpected(component));
                        }
                    }
                }
            }
            _state.OnParseRequested(this);
        }
    }
}
