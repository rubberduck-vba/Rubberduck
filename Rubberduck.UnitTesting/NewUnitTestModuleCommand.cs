using System;
using System.Linq;
using Microsoft.Vbe.Interop;
using Rubberduck.VBEditor.Extensions;

namespace Rubberduck.UnitTesting
{
    public static class NewUnitTestModuleCommand
    {
        private static readonly string TestModuleEmptyTemplate = string.Concat(
                 "'@TestModule\n"
                , "Private Assert As New Rubberduck_UnitTesting.AssertClass\n\n"
                ,"'@ModuleInitialize\n"
                ,"Public Sub ModuleInitialize()\n"
                ,"    'this method runs once per module.\n"
                ,"End Sub\n\n"
                , "'@ModuleCleanup\n"
                , "Public Sub ModuleCleanup()\n"
                , "    'this method runs once per module.\n"
                , "End Sub\n\n"
                , "'@TestInitialize\n"
                , "Public Sub TestInitialize()\n"
                , "    'this method runs before every test in the module.\n"
                , "End Sub\n\n"
                , "'@TestCleanup\n"
                , "Public Sub TestCleanup()\n"
                , "    'this method runs afer every test in the module.\n"
                , "End Sub\n\n"
            );

        private static readonly string TestModuleBaseName = "TestModule";

        public static void NewUnitTestModule(VBE vbe)
        {
            try
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

                var options = string.Concat(hasOptionExplicit ? string.Empty : "Option Explicit\n", "Option Private Module\n\n");

                module.CodeModule.AddFromString(options + TestModuleEmptyTemplate);
                module.Activate();
            }
            catch (Exception exception)
            {
            }
        }

        private static string GetNextTestModuleName(VBProject project)
        {
            var names = project.ComponentNames();
            var index = names.Count(n => n.StartsWith(TestModuleBaseName)) + 1;

            return string.Concat(TestModuleBaseName, index);
        }
    }
}
