using System;
using System.Linq;
using Microsoft.Vbe.Interop;
using Rubberduck.Settings;
using Rubberduck.UI;
using Rubberduck.VBEditor.Extensions;

namespace Rubberduck.UnitTesting
{
    public class NewUnitTestModuleCommand
    {
        private readonly VBE _vbe;
        private readonly ConfigurationLoader _configLoader;

        public NewUnitTestModuleCommand(VBE vbe, ConfigurationLoader configLoader)
        {
            _vbe = vbe;
            _configLoader = configLoader;
        }

        private readonly string _testModuleEmptyTemplate = string.Concat(
                 "'@TestModule\n"
                , "'' ", RubberduckUI.UnitTest_NewModule_UncommentLateBinding, ":\n"
                , "'Private Assert As Object\n"
                , "'' ", RubberduckUI.UnitTest_NewModule_CommentEarlyBinding, ":\n"
                , "Private Assert As New Rubberduck.AssertClass\n\n"
                , "'@ModuleInitialize\n"
                ,"Public Sub ModuleInitialize()\n"
                ,"    '", RubberduckUI.UnitTest_NewModule_RunOnce, ".\n"
                ,"    '' ", RubberduckUI.UnitTest_NewModule_UncommentLateBinding, ":\n"
                ,"    'Set Assert = CreateObject(\"Rubberduck.AssertClass\")\n"
                ,"End Sub\n\n"
                , "'@ModuleCleanup\n"
                , "Public Sub ModuleCleanup()\n"
                , "    '", RubberduckUI.UnitTest_NewModule_RunOnce, ".\n"
                , "End Sub\n\n"
                , "'@TestInitialize\n"
                , "Public Sub TestInitialize()\n"
                , "    '", RubberduckUI.UnitTest_NewModule_RunBeforeTest, ".\n"
                , "End Sub\n\n"
                , "'@TestCleanup\n"
                , "Public Sub TestCleanup()\n"
                , "    '", RubberduckUI.UnitTest_NewModule_RunAfterTest, ".\n"
                , "End Sub\n\n"
            );

        private readonly string _testModuleBaseName = "TestModule";

        public void NewUnitTestModule()
        {
            try
            {
                var project = _vbe.ActiveVBProject;
                project.EnsureReferenceToAddInLibrary();

                var module = project.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule);
                module.Name = GetNextTestModuleName(project);

                var hasOptionExplicit = false;
                if (module.CodeModule.CountOfLines > 0 && module.CodeModule.CountOfDeclarationLines > 0)
                {
                    hasOptionExplicit = module.CodeModule.Lines[1, module.CodeModule.CountOfDeclarationLines].Contains("Option Explicit");
                }

                var options = string.Concat(hasOptionExplicit ? string.Empty : "Option Explicit\n", "Option Private Module\n\n");

                module.CodeModule.AddFromString(options + _testModuleEmptyTemplate);
                module.Activate();
            }
            catch (Exception)
            {
                //can we please comment when we swallow every possible exception?
            }
        }

        private string GetNextTestModuleName(VBProject project)
        {
            var names = project.ComponentNames();
            var index = names.Count(n => n.StartsWith(_testModuleBaseName)) + 1;

            return string.Concat(_testModuleBaseName, index);
        }
    }
}
