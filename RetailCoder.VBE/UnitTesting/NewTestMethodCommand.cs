using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Reflection;
using Rubberduck.UI;
using Rubberduck.VBEditor;

namespace Rubberduck.UnitTesting
{
    public class NewTestMethodCommand
    {
        private readonly VBE _vbe;

        public NewTestMethodCommand(VBE vbe)
        {
            _vbe = vbe;
        }

        private static readonly string NamePlaceholder = "%METHODNAME%";
        private readonly string _testMethodBaseName = "TestMethod";

        private readonly string _testMethodTemplate = string.Concat(
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

        private readonly string _testMethodExpectedErrorTemplate = string.Concat(
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

        public TestMethod NewTestMethod()
        {
            if (_vbe.ActiveCodePane == null)
            {
                return null;
            }

            try
            {
                if (_vbe.ActiveCodePane.CodeModule.HasAttribute<TestModuleAttribute>())
                {
                    var module = _vbe.ActiveCodePane.CodeModule;
                    var name = GetNextTestMethodName(module.Parent);
                    var body = _testMethodTemplate.Replace(NamePlaceholder, name);
                    module.InsertLines(module.CountOfLines, body);

                    var qualifiedModuleName = new QualifiedModuleName(module.Parent);
                    return new TestMethod(new QualifiedMemberName(qualifiedModuleName, name), _vbe);
                }
            }
            catch (COMException)
            {
            }

            return null;
        }
    
        public TestMethod NewExpectedErrorTestMethod()
        {
            if (_vbe.ActiveCodePane == null)
            {
                return null;
            }

            try
            {
                if (_vbe.ActiveCodePane.CodeModule.HasAttribute<TestModuleAttribute>())
                {
                    var module = _vbe.ActiveCodePane.CodeModule;
                    var name = GetNextTestMethodName(module.Parent);
                    var body = _testMethodExpectedErrorTemplate.Replace(NamePlaceholder, name);
                    module.InsertLines(module.CountOfLines, body);

                    var qualifiedModuleName = new QualifiedModuleName(module.Parent);
                    return new TestMethod(new QualifiedMemberName(qualifiedModuleName, name), _vbe);
                }
            }
            catch (COMException)
            {
            }

            return null;
        }

        private string GetNextTestMethodName(VBComponent component)
        {
            var names = component.TestMethods().Select(test => test.QualifiedMemberName.MemberName);
            var index = names.Count(n => n.StartsWith(_testMethodBaseName)) + 1;

            return string.Concat(_testMethodBaseName, index);
        }
    }
}
