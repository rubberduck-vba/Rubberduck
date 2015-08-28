using System;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
using Rubberduck.Reflection;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;

namespace Rubberduck.UnitTesting
{
    public static class NewTestMethodCommand
    {
        private static readonly string NamePlaceholder = "%METHODNAME%";
        private static readonly string TestMethodBaseName = "TestMethod";

        private static readonly string TestMethodTemplate = string.Concat(
            "'@TestMethod\n",
            "Public Sub ", NamePlaceholder, "() 'TODO: Rename test\n",
            "    On Error GoTo TestFail\n",
            "    \n",
            "    'Arrange:\n\n",
            "    'Act:\n\n",
            "    'Assert:\n",
            "    Assert.Inconclusive\n\n",
            "TestExit:\n",
            "    Exit Sub\n",
            "TestFail:\n",
            "    Assert.Fail \"Test raised an error: #\" & Err.Number & \" - \" & Err.Description\n",
            "End Sub\n"
            );

        private static readonly string TestMethodExpectedErrorTemplate = string.Concat(
            "'@TestMethod\n",
            "Public Sub ", NamePlaceholder, "() 'TODO: Rename test\n",
            "    Const ExpectedError As Long = 0 'TODO: Change to expected error number\n",
            "    On Error GoTo TestFail\n",
            "    \n",
            "    'Arrange:\n\n",
            "    'Act:\n\n",
            "Assert:\n",
            "    Assert.Fail \"Expected error was not raised.\"\n\n",
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

        public static TestMethod NewTestMethod(VBE vbe)
        {
            if (vbe.ActiveCodePane == null)
            {
                return null;
            }

            try
            {
                if (vbe.ActiveCodePane.CodeModule.HasAttribute<TestModuleAttribute>())
                {
                    var module = vbe.ActiveCodePane.CodeModule;
                    var name = GetNextTestMethodName(module.Parent);
                    var body = TestMethodTemplate.Replace(NamePlaceholder, name);
                    module.InsertLines(module.CountOfLines, body);

                    var qualifiedModuleName = new QualifiedModuleName(module.Parent);
                    return new TestMethod(new QualifiedMemberName(qualifiedModuleName, name), vbe);
                }
            }
            catch (COMException)
            {
            }

            return null;
        }
    
        public static TestMethod NewExpectedErrorTestMethod(VBE vbe)
        {
            if (vbe.ActiveCodePane == null)
            {
                return null;
            }

            try
            {
                if (vbe.ActiveCodePane.CodeModule.HasAttribute<TestModuleAttribute>())
                {
                    var module = vbe.ActiveCodePane.CodeModule;
                    var name = GetNextTestMethodName(module.Parent);
                    var body = TestMethodExpectedErrorTemplate.Replace(NamePlaceholder, name);
                    module.InsertLines(module.CountOfLines, body);

                    var qualifiedModuleName = new QualifiedModuleName(module.Parent);
                    return new TestMethod(new QualifiedMemberName(qualifiedModuleName, name), vbe);
                }
            }
            catch (COMException)
            {
            }

            return null;
        }

        private static string GetNextTestMethodName(VBComponent component)
        {
            var names = component.TestMethods().Select(test => test.QualifiedMemberName.MemberName);
            var index = names.Count(n => n.StartsWith(TestMethodBaseName)) + 1;

            return string.Concat(TestMethodBaseName, index);
        }
    }
}