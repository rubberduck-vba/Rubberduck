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
    public class AddTestMethodCommand : ComCommandBase
    {
        private readonly IVBE _vbe;
        private readonly RubberduckParserState _state;
        private readonly ITestCodeGenerator _codeGenerator;

        public AddTestMethodCommand(
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
            if (_state.Status != ParserState.Ready)
            {
                return false;
            }

            using (var activePane = _vbe.ActiveCodePane)
            {
                if (activePane?.IsWrappingNullReference ?? true)
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
                using( var activeCodePane = _vbe.ActiveCodePane)
                {
                    using( var activePaneCodeModule = activeCodePane.CodeModule)
                    {
                        return testModules.Any(a => _state.ProjectsProvider.Component(a.QualifiedModuleName).HasEqualCodeModule(activePaneCodeModule));
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

                using (var module = pane.CodeModule)
                {
                    var declaration = _state.GetTestModules()
                        .FirstOrDefault(f => _state.ProjectsProvider.Component(f.QualifiedModuleName).HasEqualCodeModule(module));

                    if (declaration == null)
                    {
                        return;
                    }

                    using (var component = module.Parent)
                    {
                        module.InsertLines(module.CountOfLines, _codeGenerator.GetNewTestMethodCode(component));
                    }
                }
            }
            _state.OnParseRequested(this);
        }
    }
}
