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
    public class AddTestMethodExpectedErrorCommand : CommandBase
    {
        private readonly IVBE _vbe;
        private readonly RubberduckParserState _state;

        public AddTestMethodExpectedErrorCommand(IVBE vbe, RubberduckParserState state) : base(LogManager.GetCurrentClassLogger())
        {
            _vbe = vbe;
            _state = state;
        }

        public const string NamePlaceholder = "%METHODNAME%";
        private const string TestMethodBaseName = "TestMethod";

        public static readonly string TestMethodExpectedErrorTemplate = string.Concat(
            "'@TestMethod\r\n",
            "Public Sub ", NamePlaceholder, "() 'TODO ", TestExplorer.UnitTest_NewMethod_Rename, "\r\n",
            "    Const ExpectedError As Long = 0 'TODO ", TestExplorer.UnitTest_NewMethod_ChangeErrorNo, "\r\n",
            "    On Error GoTo TestFail\r\n",
            "    \r\n",
            "    'Arrange:\r\n\r\n",
            "    'Act:\r\n\r\n",
            "Assert:\r\n",
            "    Assert.Fail \"", TestExplorer.UnitTest_NewMethod_ErrorNotRaised, ".\"\r\n\r\n",
            "TestExit:\r\n",
            "    Exit Sub\r\n",
            "TestFail:\r\n",
            "    If Err.Number = ExpectedError Then\r\n",
            "        Resume TestExit\r\n",
            "    Else\r\n",
            "        Resume Assert\r\n",
            "    End If\r\n",
            "End Sub\r\n"
            );

        protected override bool EvaluateCanExecute(object parameter)
        {
            using (var pane = _vbe.ActiveCodePane)
            {
                if (_state.Status != ParserState.Ready || pane.IsWrappingNullReference)
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
                if (pane.IsWrappingNullReference)
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
                        string name;
                        using (var component = activeModule.Parent)
                        {
                            name = GetNextTestMethodName(component);
                        }
                        var body = TestMethodExpectedErrorTemplate.Replace(NamePlaceholder, name);
                        activeModule.InsertLines(activeModule.CountOfLines, body);
                        
                    }
                }
            }
            _state.OnParseRequested(this);
        }

        [Obsolete("Duplicates AddTestMethodCommand#GetNextTestMethodName, should be centrally solved in UnitTesting assembly instead")]
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
