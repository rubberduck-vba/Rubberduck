using System;
using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.Parsing.Annotations.Concrete;
using Rubberduck.Parsing.Rewriter;
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
    public class AddTestMethodBase : ComCommandBase
    {
        private readonly IVBE _vbe;
        private readonly RubberduckParserState _state;
        private readonly IRewritingManager _rewritingManager;
        private readonly ITestCodeGenerator _testCodeGenerator;

        public AddTestMethodBase(
            IVBE vbe,
            RubberduckParserState state,
            IRewritingManager rewritingManager,
            ITestCodeGenerator codeGenerator,
            IVbeEvents vbeEvents)
            : base(vbeEvents)
        {
            _vbe = vbe;
            _state = state;
            _testCodeGenerator = codeGenerator;
            _rewritingManager = rewritingManager;

            MethodGenerator = _testCodeGenerator.GetNewTestMethodCode;

            AddToCanExecuteEvaluation(SpecialEvaluateCanExecute);
        }

        protected Func<IVBComponent, string> MethodGenerator { set; get; }

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
                using (var activeCodePane = _vbe.ActiveCodePane)
                {
                    using (var activePaneCodeModule = activeCodePane.CodeModule)
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
            (Declaration targetModule, string testMethodCode) = GetTargetModuleAndTestMethodCode(_vbe, _state, MethodGenerator);

            if (targetModule != null)
            {
                var rewriteSession = _rewritingManager.CheckOutCodePaneSession();
                var rewriter = rewriteSession.CheckOutModuleRewriter(targetModule.QualifiedModuleName);
                rewriter.InsertAfter(rewriter.TokenStream.Size, $"{Environment.NewLine}{Environment.NewLine}{testMethodCode}");

                rewriteSession.TryRewrite();

                _state.OnParseRequested(this);
            }
        }

        private static  (Declaration targetModule, string testMethodContent) GetTargetModuleAndTestMethodCode(IVBE vbe, RubberduckParserState state, Func<IVBComponent, string> methodGenerator)
        {
            (Declaration, string) defaultResult = (null, string.Empty);

            using (var pane = vbe.ActiveCodePane)
            {
                if (pane?.IsWrappingNullReference ?? true)
                {
                    return defaultResult;
                }

                using (var module = pane.CodeModule)
                {
                    var declaration = state.GetTestModules()
                        .FirstOrDefault(f => state.ProjectsProvider.Component(f.QualifiedModuleName).HasEqualCodeModule(module));

                    if (declaration == null)
                    {
                        return defaultResult;
                    }

                    var testMethodCode = string.Empty;
                    using (var component = module.Parent)
                    {
                        testMethodCode = methodGenerator(component);
                        return (declaration, testMethodCode);
                    }
                }
            }
        }
    }
}
