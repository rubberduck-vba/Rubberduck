using System;
using System.Linq;
using Microsoft.Vbe.Interop;
using Rubberduck.VBEditor.Extensions;

namespace Rubberduck.UnitTesting
{
    public static class NewUnitTestModuleCommand
    {
        private static readonly string TestModuleEmptyTemplate = String.Concat(
                 "'@TestModule\n"
                , "'' ", Rubberduck.UI.RubberduckUI.UnitTest_NewModule_UncommentLateBinding, ":\n"
                , "'Private Assert As Object\n"
                , "'' ", Rubberduck.UI.RubberduckUI.UnitTest_NewModule_CommentEarlyBinding, ":\n"
                , "Private Assert As New Rubberduck.AssertClass\n\n"
                , "'@ModuleInitialize\n"
                ,"Public Sub ModuleInitialize()\n"
                ,"    '", Rubberduck.UI.RubberduckUI.UnitTest_NewModule_RunOnce, ".\n"
                ,"    '' ", Rubberduck.UI.RubberduckUI.UnitTest_NewModule_UncommentLateBinding, ":\n"
                ,"    'Set Assert = CreateObject(\"Rubberduck.AssertClass\")\n"
                ,"End Sub\n\n"
                , "'@ModuleCleanup\n"
                , "Public Sub ModuleCleanup()\n"
                , "    '", Rubberduck.UI.RubberduckUI.UnitTest_NewModule_RunOnce, ".\n"
                , "End Sub\n\n"
                , "'@TestInitialize\n"
                , "Public Sub TestInitialize()\n"
                , "    '", Rubberduck.UI.RubberduckUI.UnitTest_NewModule_RunBeforeTest, ".\n"
                , "End Sub\n\n"
                , "'@TestCleanup\n"
                , "Public Sub TestCleanup()\n"
                , "    '", Rubberduck.UI.RubberduckUI.UnitTest_NewModule_RunAfterTest, ".\n"
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
                //can we please comment when we swallow every possible exception?
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
