using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using Rubberduck.Reflection;
using System;

namespace Rubberduck.UnitTesting
{
    internal static class NewUnitTestModuleCommand
    {
        private static readonly string TestModuleEmptyTemplate = string.Concat(
            "'@TestModule\n",
            "Private Assert As New Rubberduck.AssertClass\n\n"
            );

        private static readonly string TestModuleTemplate = string.Concat(
            TestModuleEmptyTemplate,
            "'test explorer searches in standard code modules, for all\n",
            "'public parameterless procedures with a name that starts with \"Test\".\n",
            "Public Sub TestMethod1()\n",
            "    Assert.Inconclusive\n",
            "End Sub\n\n",
            "'...or public parameterless procedures with a @TestMethod marker:\n",
            "'@TestMethod\n",
            "Public Sub TestMethod2()\n",
            "    Assert.Inconclusive\n",
            "End Sub\n\n",
            "'test methods that make no assertions will succeed:\n",
            "'@TestMethod\n",
            "Public Sub TestMethod3()\n",
            "End Sub\n\n"
            );

        private static readonly string TestModuleBaseName = "TestModule";

        public static void NewUnitTestModule(VBE vbe)
        {
            var project = vbe.ActiveVBProject;
            project.EnsureReferenceToAddInLibrary();

            var module = project.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule);
            module.Name = GetNextTestModuleName(project);

            var hasOptionExplicit = false;
            if (module.CodeModule.CountOfLines > 0 && module.CodeModule.CountOfDeclarationLines > 0)
            {
                hasOptionExplicit = module.CodeModule.Lines[1, module.CodeModule.CountOfDeclarationLines].Contains("Option Explicit");
            }
            module.CodeModule.AddFromString((hasOptionExplicit ? string.Empty : "Option Explicit\n") + TestModuleEmptyTemplate);
            module.Activate();
        }

        public static void NewUnitTestModuleTemplate(VBE vbe)
        {
            var project = vbe.ActiveVBProject;
            project.EnsureReferenceToAddInLibrary();

            var module = project.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule);
            module.Name = GetNextTestModuleName(project);
            module.CodeModule.AddFromString(TestModuleTemplate);
            module.Activate();
        }

        private static string GetNextTestModuleName(VBProject project)
        {
            var names = project.ComponentNames();
            var index = names.Count(n => n.StartsWith(TestModuleBaseName)) + 1;

            return string.Concat(TestModuleBaseName, index);
        }
    }

    internal static class NewTestMethodCommand
    {
        private static readonly string NamePlaceholder = "%METHODNAME%";
        private static readonly string TestMethodBaseName = "TestMethod";

        private static readonly string TestMethodTemplate = string.Concat(
            "'@TestMethod\n",
            "Public Sub ", NamePlaceholder, "() 'TODO: Rename test\n",
            "    On Error GoTo TestFail\n",
            "    \n",
            "Arrange:\n\n",
            "Act:\n\n",
            "Assert:\n",
            "    Assert.Inconclusive\n\n",
            "TestExit:\n",
            "    Exit Sub\n",
            "TestFail:\n",
            "    If Err.Number <> 0 Then\n",
            "        Assert.Fail \"Test raised an error: #\" & Err.Number & \" - \" & Err.Description\n",
            "    Else\n",
            "        Resume TestExit\n",
            "    End If\n",
            "End Sub\n"
            );

        private static readonly string TestMethodExpectedErrorTemplate = string.Concat(
            "'@TestMethod\n",
            "Public Sub ", NamePlaceholder, "() 'TODO: Rename test\n",
            "    Const ExpectedError As Long = 0 'TODO: Change to expected error number\n",
            "    On Error GoTo TestFail\n",
            "    \n",
            "Arrange:\n\n",
            "Act:\n\n",
            "Assert:\n",
            "    Assert.Fail \"Expected error was not raised.\"\n\n",
            "TestExit:\n",
            "    Exit Sub\n",
            "TestFail:\n",
            "    If Err.Number <> ExpectedError Then\n",
            "        Resume Assert\n",
            "    Else\n",
            "        Resume TestExit\n",
            "    End If\n",
            "End Sub\n"
            );

        public static void NewTestMethod(VBE vbe)
        {
            if (vbe.ActiveCodePane.CodeModule.HasAttribute<TestModuleAttribute>())
            {
                var module = vbe.ActiveCodePane.CodeModule;
                var name = GetNextTestMethodName(module.Parent);
                var method = TestMethodTemplate.Replace(NamePlaceholder, name);
                module.InsertLines(module.CountOfLines, method);
            }
        }

        public static void NewExpectedErrorTestMethod(VBE vbe)
        {
            if (vbe.ActiveCodePane.CodeModule.HasAttribute<TestModuleAttribute>())
            {
                var module = vbe.ActiveCodePane.CodeModule;
                var name = GetNextTestMethodName(module.Parent);
                var method = TestMethodExpectedErrorTemplate.Replace(NamePlaceholder, name);
                module.InsertLines(module.CountOfLines, method);
            }
        }
        private static string GetNextTestMethodName(VBComponent component)
        {
            var names = component.TestMethods().Select(test => test.MethodName);
            var index = names.Count(n => n.StartsWith(TestMethodBaseName)) + 1;

            return string.Concat(TestMethodBaseName, index);
        }

    }
}
