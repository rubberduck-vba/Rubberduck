using System.Linq;
using System.Runtime.InteropServices;
using NLog;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.UnitTesting;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.VBA;

namespace Rubberduck.UI.Command
{
    /// <summary>
    /// A command that adds a new test module to the active VBAProject.
    /// </summary>
    [ComVisible(false)]
    public class AddTestModuleCommand : CommandBase
    {
        private readonly IVBE _vbe;
        private readonly RubberduckParserState _state;
        private readonly IGeneralConfigService _configLoader;

        public AddTestModuleCommand(IVBE vbe, RubberduckParserState state, IGeneralConfigService configLoader)
            : base(LogManager.GetCurrentClassLogger())
        {
            _vbe = vbe;
            _state = state;
            _configLoader = configLoader;
        }

        private const string FolderAnnotation = "'@Folder(\"Tests\")\r\n";
        private const string ModuleLateBinding = "Private Assert As Object\r\n";
        private const string ModuleEarlyBinding = "Private Assert As New Rubberduck.{0}AssertClass\r\n";

        private const string TestModuleEmptyTemplate = "'@TestModule\r\n{0}\r\n{1}\r\n";

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

        private string GetTestModule(IUnitTestSettings settings)
        {
            var assertClass = settings.AssertMode == AssertMode.StrictAssert ? string.Empty : "Permissive";
            var moduleBinding = settings.BindingMode == BindingMode.EarlyBinding
                ? string.Format(ModuleEarlyBinding, assertClass)
                : ModuleLateBinding;

            var formattedModuleTemplate = string.Format(TestModuleEmptyTemplate, FolderAnnotation, moduleBinding);

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

        private IVBProject GetProject()
        {
            var activeProject = _vbe.ActiveVBProject;
            if (!activeProject.IsWrappingNullReference)
            {
                return activeProject;
            }

            var projects = _vbe.VBProjects;
            {
                return projects.Count == 1 
                    ? projects[1]
                    : new VBProject(null);
            }
        }

        protected override bool CanExecuteImpl(object parameter)
        {
            return !GetProject().IsWrappingNullReference && _vbe.HostSupportsUnitTests();
        }

        protected override void ExecuteImpl(object parameter)
        {
            var project = parameter as IVBProject ?? GetProject();
            if (project.IsWrappingNullReference)
            {
                return;
            }

            var settings = _configLoader.LoadConfiguration().UserSettings.UnitTestSettings;

            if (settings.BindingMode == BindingMode.EarlyBinding)
            {
                project.EnsureReferenceToAddInLibrary();
            }

            var component = project.VBComponents.Add(ComponentType.StandardModule);
            var module = component.CodeModule;
            component.Name = GetNextTestModuleName(project);

            var hasOptionExplicit = false;
            if (module.CountOfLines > 0 && module.CountOfDeclarationLines > 0)
            {
                hasOptionExplicit = module.GetLines(1, module.CountOfDeclarationLines).Contains("Option Explicit");
            }

            var options = string.Concat(hasOptionExplicit ? string.Empty : "Option Explicit\r\n",
                "Option Private Module\r\n\r\n");

            var defaultTestMethod = string.Empty;
            if (settings.DefaultTestStubInNewModule)
            {
                defaultTestMethod = AddTestMethodCommand.TestMethodTemplate.Replace(
                    AddTestMethodCommand.NamePlaceholder, "TestMethod1");
            }

            module.AddFromString(options + GetTestModule(settings) + defaultTestMethod);
            component.Activate();
            _state.OnParseRequested(this, component);
        }

        private string GetNextTestModuleName(IVBProject project)
        {
            var names = project.ComponentNames();
            var index = names.Count(n => n.StartsWith(TestModuleBaseName)) + 1;

            return string.Concat(TestModuleBaseName, index);
        }
    }
}
