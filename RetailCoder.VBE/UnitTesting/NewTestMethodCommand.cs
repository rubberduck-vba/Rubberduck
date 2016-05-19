using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;

namespace Rubberduck.UnitTesting
{
    public class NewTestMethodCommand
    {
        private readonly VBE _vbe;
        private readonly RubberduckParserState _state;

        public NewTestMethodCommand(VBE vbe, RubberduckParserState state)
        {
            _vbe = vbe;
            _state = state;
        }

        public const string NamePlaceholder = "%METHODNAME%";
        private const string TestMethodBaseName = "TestMethod";

        public static readonly string TestMethodTemplate = string.Concat(
            "'@TestMethod\n",
            "Public Sub ", NamePlaceholder, "() 'TODO ", RubberduckUI.UnitTest_NewMethod_Rename, "\n",
            "    On Error GoTo TestFail\n",
            "    \n",
            "    'Arrange:\n\n",
            "    'Act:\n\n",
            "    'Assert:\n",
            "    Assert.Inconclusive\n\n",
            "TestExit:\n",
            "    Exit Sub\n",
            "TestFail:\n",
            "    Assert.Fail \"", RubberduckUI.UnitTest_NewMethod_RaisedTestError, ": #\" & Err.Number & \" - \" & Err.Description\n",
            "End Sub\n"
            );

        public static readonly string TestMethodExpectedErrorTemplate = string.Concat(
            "'@TestMethod\n",
            "Public Sub ", NamePlaceholder, "() 'TODO ", RubberduckUI.UnitTest_NewMethod_Rename, "\n",
            "    Const ExpectedError As Long = 0 'TODO ", RubberduckUI.UnitTest_NewMethod_ChangeErrorNo, "\n",
            "    On Error GoTo TestFail\n",
            "    \n",
            "    'Arrange:\n\n",
            "    'Act:\n\n",
            "Assert:\n",
            "    Assert.Fail \"", RubberduckUI.UnitTest_NewMethod_ErrorNotRaised, ".\"\n\n",
            "TestExit:\n",
            "    Exit Sub\n",
            "TestFail:\n",
            "    If Err.Number = ExpectedError Then\n",
            "        Resume TestExit\n",
            "    Else\n",
            "        Resume Assert\n",
            "    End If\n",
            "End Sub\n"
            );

        public void NewTestMethod()
        {
            if (_vbe.ActiveCodePane == null)
            {
                return;
            }

            try
            {
                var declaration = _state.AllUserDeclarations.First(f =>
                            f.DeclarationType == Parsing.Symbols.DeclarationType.ProceduralModule &&
                            f.QualifiedName.QualifiedModuleName.Component.CodeModule == _vbe.ActiveCodePane.CodeModule);

                if (declaration.Annotations.Any(a => a.AnnotationType == AnnotationType.TestModule))
                {
                    var module = _vbe.ActiveCodePane.CodeModule;
                    var name = GetNextTestMethodName(module.Parent);
                    var body = TestMethodTemplate.Replace(NamePlaceholder, name);
                    module.InsertLines(module.CountOfLines, body);
                }
            }
            catch (COMException)
            {
            }

            _state.OnParseRequested(this, _vbe.SelectedVBComponent);
        }
    
        public void NewExpectedErrorTestMethod()
        {
            if (_vbe.ActiveCodePane == null)
            {
                return;
            }

            try
            {
                var declaration = _state.AllUserDeclarations.First(f =>
                            f.DeclarationType == Parsing.Symbols.DeclarationType.ProceduralModule &&
                            f.QualifiedName.QualifiedModuleName.Component.CodeModule == _vbe.ActiveCodePane.CodeModule);

                if (declaration.Annotations.Any(a => a.AnnotationType == AnnotationType.TestModule))
                {
                    var module = _vbe.ActiveCodePane.CodeModule;
                    var name = GetNextTestMethodName(module.Parent);
                    var body = TestMethodExpectedErrorTemplate.Replace(NamePlaceholder, name);
                    module.InsertLines(module.CountOfLines, body);
                }
            }
            catch (COMException)
            {
            }

            _state.OnParseRequested(this, _vbe.SelectedVBComponent);
        }

        private string GetNextTestMethodName(VBComponent component)
        {
            var names = component.GetTests(_vbe, _state).Select(test => test.QualifiedMemberName.MemberName);
            var index = names.Count(n => n.StartsWith(TestMethodBaseName)) + 1;

            return string.Concat(TestMethodBaseName, index);
        }
    }
}
