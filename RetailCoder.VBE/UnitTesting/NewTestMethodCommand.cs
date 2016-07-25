using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
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
            "'@TestMethod\r\n",
            "Public Sub ", NamePlaceholder, "() 'TODO ", RubberduckUI.UnitTest_NewMethod_Rename, "\r\n",
            "    On Error GoTo TestFail\r\n",
            "    \r\n",
            "    'Arrange:\r\n\r\n",
            "    'Act:\r\n\r\n",
            "    'Assert:\r\n",
            "    Assert.Inconclusive\r\n\r\n",
            "TestExit:\r\n",
            "    Exit Sub\r\n",
            "TestFail:\r\n",
            "    Assert.Fail \"", RubberduckUI.UnitTest_NewMethod_RaisedTestError, ": #\" & Err.Number & \" - \" & Err.Description\r\n",
            "End Sub\r\n"
            );

        public static readonly string TestMethodExpectedErrorTemplate = string.Concat(
            "'@TestMethod\r\n",
            "Public Sub ", NamePlaceholder, "() 'TODO ", RubberduckUI.UnitTest_NewMethod_Rename, "\r\n",
            "    Const ExpectedError As Long = 0 'TODO ", RubberduckUI.UnitTest_NewMethod_ChangeErrorNo, "\r\n",
            "    On Error GoTo TestFail\r\n",
            "    \r\n",
            "    'Arrange:\r\n\r\n",
            "    'Act:\r\n\r\n",
            "Assert:\r\n",
            "    Assert.Fail \"", RubberduckUI.UnitTest_NewMethod_ErrorNotRaised, ".\"\r\n\r\n",
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

        public void NewTestMethod()
        {
            if (_vbe.ActiveCodePane == null)
            {
                return;
            }

            try
            {
                var declaration = _state.GetTestModules().FirstOrDefault(f =>
                            f.QualifiedName.QualifiedModuleName.Component.CodeModule == _vbe.ActiveCodePane.CodeModule);

                if (declaration != null)
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
                var declaration = _state.GetTestModules().FirstOrDefault(f =>
                            f.QualifiedName.QualifiedModuleName.Component.CodeModule == _vbe.ActiveCodePane.CodeModule);

                if (declaration != null)
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
            var names = component.GetTests(_vbe, _state).Select(test => test.Declaration.IdentifierName);
            var index = names.Count(n => n.StartsWith(TestMethodBaseName)) + 1;

            return string.Concat(TestMethodBaseName, index);
        }
    }
}
