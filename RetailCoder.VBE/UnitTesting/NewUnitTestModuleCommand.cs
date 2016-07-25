using System;
using System.Linq;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.UI;
using Rubberduck.VBEditor.Extensions;

namespace Rubberduck.UnitTesting
{
    public class NewUnitTestModuleCommand
    {
        private readonly RubberduckParserState _state;
        private readonly ConfigurationLoader _configLoader;

        public NewUnitTestModuleCommand(RubberduckParserState state, ConfigurationLoader configLoader)
        {
            _state = state;
            _configLoader = configLoader;
        }

        private const string ModuleLateBinding = "Private Assert As Object\r\n";
        private const string ModuleEarlyBinding = "Private Assert As New Rubberduck.{0}AssertClass\r\n";

        private const string TestModuleEmptyTemplate = "'@TestModule\r\n{0}\r\n";

        private const string ModuleInitLateBinding = "Set Assert = CreateObject(\"Rubberduck.{0}AssertClass\")\r\n";
        private readonly string _moduleInit = string.Concat(
            "'@ModuleInitialize\r\n"
            , "Public Sub ModuleInitialize()\r\n"
            , "    '", RubberduckUI.UnitTest_NewModule_RunOnce, ".\r\n    {0}\r\n"
            , "End Sub\r\n\r\n"
            , "'@ModuleCleanup\r\n"
            , "Public Sub ModuleCleanup()\r\n"
            , "    '", RubberduckUI.UnitTest_NewModule_RunOnce, ".\r\n"
            , "End Sub\r\n\r\n"
        );

        private readonly string _methodInit = string.Concat(
            "'@TestInitialize\r\n"
            , "Public Sub TestInitialize()\r\n"
            , "    '", RubberduckUI.UnitTest_NewModule_RunBeforeTest, ".\r\n"
            , "End Sub\r\n\r\n"
            , "'@TestCleanup\r\n"
            , "Public Sub TestCleanup()\r\n"
            , "    '", RubberduckUI.UnitTest_NewModule_RunAfterTest, ".\r\n"
            , "End Sub\r\n\r\n"
        );

        private const string TestModuleBaseName = "TestModule";

        private string GetTestModule(UnitTestSettings settings)
        {
            var assertClass = settings.AssertMode == AssertMode.StrictAssert ? string.Empty : "Permissive";
            var moduleBinding = settings.BindingMode == BindingMode.EarlyBinding
                ? string.Format(ModuleEarlyBinding, assertClass)
                : ModuleLateBinding;

            var formattedModuleTemplate = string.Format(TestModuleEmptyTemplate, moduleBinding);

            if (settings.ModuleInit)
            {
                var lateBindingString = string.Format(ModuleInitLateBinding,
                    settings.AssertMode == AssertMode.StrictAssert ? string.Empty : "Permissive");

                formattedModuleTemplate += string.Format(_moduleInit, settings.BindingMode == BindingMode.EarlyBinding ? string.Empty : lateBindingString);
            }

            if (settings.MethodInit)
            {
                formattedModuleTemplate += _methodInit;
            }

            return formattedModuleTemplate;
        }

        public void NewUnitTestModule(VBProject project)
        {
            var settings = _configLoader.LoadConfiguration().UserSettings.UnitTestSettings;
            VBComponent component;
            
            try
            {
                if (settings.BindingMode == BindingMode.EarlyBinding)
                {
                    project.EnsureReferenceToAddInLibrary();
                }

                component = project.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule);
                component.Name = GetNextTestModuleName(project);

                var hasOptionExplicit = false;
                if (component.CodeModule.CountOfLines > 0 && component.CodeModule.CountOfDeclarationLines > 0)
                {
                    hasOptionExplicit = component.CodeModule.Lines[1, component.CodeModule.CountOfDeclarationLines].Contains("Option Explicit");
                }

                var options = string.Concat(hasOptionExplicit ? string.Empty : "Option Explicit\r\n", "Option Private Module\r\n\r\n");

                var defaultTestMethod = string.Empty;
                if (settings.DefaultTestStubInNewModule)
                {
                    defaultTestMethod = NewTestMethodCommand.TestMethodTemplate.Replace(
                        NewTestMethodCommand.NamePlaceholder, "TestMethod1");
                }
                
                component.CodeModule.AddFromString(options + GetTestModule(settings) + defaultTestMethod);
                component.Activate();
            }
            catch (Exception)
            {
                //can we please comment when we swallow every possible exception?
                return;
            }

            _state.OnParseRequested(this, component);
        }

        private string GetNextTestModuleName(VBProject project)
        {
            var names = project.ComponentNames();
            var index = names.Count(n => n.StartsWith(TestModuleBaseName)) + 1;

            return string.Concat(TestModuleBaseName, index);
        }
    }
}
