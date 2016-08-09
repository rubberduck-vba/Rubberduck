using System;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using NLog;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.UnitTesting;
using Rubberduck.VBEditor.Extensions;

namespace Rubberduck.UI.Command
{
    /// <summary>
    /// A command that adds a new test module to the active VBAProject.
    /// </summary>
    [ComVisible(false)]
    public class AddTestModuleCommand : CommandBase
    {
        private readonly VBE _vbe;
        private readonly RubberduckParserState _state;
        private readonly IGeneralConfigService _configLoader;

        public AddTestModuleCommand(VBE vbe, RubberduckParserState state, IGeneralConfigService configLoader)
            : base(LogManager.GetCurrentClassLogger())
        {
            _vbe = vbe;
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

        private VBProject GetProject()
        {
            return _vbe.ActiveVBProject ?? (_vbe.VBProjects.Count == 1 ? _vbe.VBProjects.Item(1) : null);
        }

        protected override bool CanExecuteImpl(object parameter)
        {
            return GetProject() != null &&
                _vbe.HostSupportsUnitTests();
        }
        
        protected override void ExecuteImpl(object parameter)
        {
            var project = parameter as VBProject ?? GetProject();
            if (project == null) { return; }

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
                    defaultTestMethod = AddTestMethodCommand.TestMethodTemplate.Replace(
                        AddTestMethodCommand.NamePlaceholder, "TestMethod1");
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
