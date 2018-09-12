using System.Collections.Generic;
using System;
using System.Linq;
using System.Runtime.InteropServices;
using NLog;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources.UnitTesting;
using Rubberduck.UI.Command;
using Rubberduck.UnitTesting;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.UnitTesting.Commands
{
    /// <summary>
    /// A command that adds a new test method stub to the active code pane.
    /// </summary>
    [ComVisible(false)]
    public class AddTestMethodCommand : CommandBase
    {
        private readonly IVBE _vbe;
        private readonly RubberduckParserState _state;

        public AddTestMethodCommand(IVBE vbe, RubberduckParserState state) : base(LogManager.GetCurrentClassLogger())
        {
            _vbe = vbe;
            _state = state;
        }

        public const string NamePlaceholder = "%METHODNAME%";
        private const string TestMethodBaseName = "TestMethod";

        public static readonly string TestMethodTemplate = string.Concat(
            "'@TestMethod\r\n",
            "Public Sub ", NamePlaceholder, "() 'TODO ", TestExplorer.UnitTest_NewMethod_Rename, "\r\n",
            "    On Error GoTo TestFail\r\n",
            "    \r\n",
            "    'Arrange:\r\n\r\n",
            "    'Act:\r\n\r\n",
            "    'Assert:\r\n",
            "    Assert.Inconclusive\r\n\r\n",
            "TestExit:\r\n",
            "    Exit Sub\r\n",
            "TestFail:\r\n",
            "    Assert.Fail \"", TestExplorer.UnitTest_NewMethod_RaisedTestError, ": #\" & Err.Number & \" - \" & Err.Description\r\n",
            "End Sub\r\n"
            );

        protected override bool EvaluateCanExecute(object parameter)
        {
            if (_state.Status != ParserState.Ready)
            {
                return false;
            }

            using (var activePane = _vbe.ActiveCodePane)
            {
                if (activePane == null || activePane.IsWrappingNullReference)
                {
                    return false;
                }
            }

            var testModules = _state.AllUserDeclarations.Where(d =>
                        d.DeclarationType == DeclarationType.ProceduralModule &&
                        d.Annotations.Any(a => a.AnnotationType == AnnotationType.TestModule));

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
                if (pane.IsWrappingNullReference)
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

                    string name;
                    using (var component = module.Parent)
                    {
                        name = GetNextTestMethodName(component);
                    }
                    var body = TestMethodTemplate.Replace(NamePlaceholder, name);
                    module.InsertLines(module.CountOfLines, body);
                }
            }
            _state.OnParseRequested(this);
        }

        [Obsolete("Duplicates AddTestMethodExpectedErrorCommand#GetNextTestMethodName, should be centrally solved in UnitTesting assembly instead")]
        private string GetNextTestMethodName(IVBComponent component)
        {
            var names = new HashSet<string>(_state.DeclarationFinder.Members(component.QualifiedModuleName)
                .Select(test => test.IdentifierName).Where(decl => decl.StartsWith(TestMethodBaseName)));

            var index = 1;
            while (names.Contains($"{TestMethodBaseName}{index}"))
            {
                index++;
            }

            return $"{TestMethodBaseName}{index}";
        }
    }
}
