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
        private readonly VBE _vbe;
        private readonly RubberduckParserState _state;
        private readonly ConfigurationLoader _configLoader;

        public NewUnitTestModuleCommand(VBE vbe, RubberduckParserState state, ConfigurationLoader configLoader)
        {
            _vbe = vbe;
            _state = state;
            _configLoader = configLoader;
        }

        private const string ModuleLateBinding = "Private Assert As Object\n";
        private const string ModuleEarlyBinding = "Private Assert As New Rubberduck.{0}AssertClass\n";

        private const string TestModuleEmptyTemplate = "'@TestModule\n{0}\n";

        private const string ModuleInitLateBinding = "Set Assert = CreateObject(\"Rubberduck.{0}AssertClass\")\n";
        private readonly string _moduleInit = string.Concat(
            "'@ModuleInitialize\n"
            , "Public Sub ModuleInitialize()\n"
            , "    '", RubberduckUI.UnitTest_NewModule_RunOnce, ".\n{0}\n"
            , "End Sub\n\n"
            , "'@ModuleCleanup\n"
            , "Public Sub ModuleCleanup()\n"
            , "    '", RubberduckUI.UnitTest_NewModule_RunOnce, ".\n"
            , "End Sub\n\n"
        );

        private readonly string _methodInit = string.Concat(
            "'@TestInitialize\n"
            , "Public Sub TestInitialize()\n"
            , "    '", RubberduckUI.UnitTest_NewModule_RunBeforeTest, ".\n"
            , "End Sub\n\n"
            , "'@TestCleanup\n"
            , "Public Sub TestCleanup()\n"
            , "    '", RubberduckUI.UnitTest_NewModule_RunAfterTest, ".\n"
            , "End Sub\n\n"
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
                project.EnsureReferenceToAddInLibrary();

                component = project.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule);
                component.Name = GetNextTestModuleName(project);

                var hasOptionExplicit = false;
                if (component.CodeModule.CountOfLines > 0 && component.CodeModule.CountOfDeclarationLines > 0)
                {
                    hasOptionExplicit = component.CodeModule.Lines[1, component.CodeModule.CountOfDeclarationLines].Contains("Option Explicit");
                }

                var options = string.Concat(hasOptionExplicit ? string.Empty : "Option Explicit\n", "Option Private Module\n\n");

                component.CodeModule.AddFromString(options + GetTestModule(settings));
                component.Activate();
            }
            catch (Exception)
            {
                //can we please comment when we swallow every possible exception?
                return;
            }

            _state.StateChanged += (sender, args) =>
            {
                if (args.State == ParserState.Ready && settings.DefaultTestStubInNewModule)
                {
                    var newTestMethodCommand = new NewTestMethodCommand(_vbe, _state);
                    newTestMethodCommand.NewTestMethod();
                }
            };

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
