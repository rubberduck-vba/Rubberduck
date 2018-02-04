using System.Collections.Generic;
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
using System.Text;
using Rubberduck.Parsing.Symbols;

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

        private readonly string _testModuleEmptyTemplate = new StringBuilder()
            .AppendLine("'@TestModule")
            .AppendLine("'@Folder(\"Tests\")")
            .AppendLine()
            .AppendLine("{0}")
            .AppendLine("{1}")
            .AppendLine()
            .ToString();

        private readonly string _moduleInit = new StringBuilder()
            .AppendLine("'@ModuleInitialize")
            .AppendLine("Public Sub ModuleInitialize()")
            .AppendLine($"    '{RubberduckUI.UnitTest_NewModule_RunOnce}.")
            .AppendLine("    {0}")
            .AppendLine("    {1}")
            .AppendLine("End Sub")
            .AppendLine()
            .AppendLine("'@ModuleCleanup")
            .AppendLine("Public Sub ModuleCleanup()")
            .AppendLine($"    '{RubberduckUI.UnitTest_NewModule_RunOnce}.")
            .AppendLine("    Set Assert = Nothing")
            .AppendLine("    Set Fakes = Nothing")
            .AppendLine("End Sub")
            .AppendLine()
            .ToString();

        private readonly string _methodInit = new StringBuilder()
            .AppendLine("'@TestInitialize")
            .AppendLine("Public Sub TestInitialize()")
            .AppendLine($"    '{RubberduckUI.UnitTest_NewModule_RunBeforeTest}.")
            .AppendLine("End Sub")
            .AppendLine()
            .AppendLine("'@TestCleanup")
            .AppendLine("Public Sub TestCleanup()")
            .AppendLine($"    '{RubberduckUI.UnitTest_NewModule_RunAfterTest}.")
            .AppendLine("End Sub")
            .AppendLine()
            .ToString();

        private const string FakesFieldDeclarationFormat = "Private Fakes As {0}";
        private const string AssertFieldDeclarationFormat = "Private Assert As {0}";

        private const string TestModuleBaseName = "TestModule";

        private string GetTestModule(IUnitTestSettings settings)
        {
            var assertType = string.Format("Rubberduck.{0}AssertClass", settings.AssertMode == AssertMode.StrictAssert ? string.Empty : "Permissive");
            var assertDeclaredAs = DeclarationFormatFor(AssertFieldDeclarationFormat, assertType, settings);

            const string fakesType = "Rubberduck.FakesProvider";
            var fakesDeclaredAs = DeclarationFormatFor(FakesFieldDeclarationFormat, fakesType, settings); 

            var formattedModuleTemplate = string.Format(_testModuleEmptyTemplate, assertDeclaredAs, fakesDeclaredAs);

            if (settings.ModuleInit)
            {
                var assertBinding = InstantiationFormatFor(assertType, settings);
                var assertSetAs = $"Set Assert = {assertBinding}";

                var fakesBinding = InstantiationFormatFor(fakesType, settings);
                var fakesSetAs = $"Set Fakes = {fakesBinding}";

                formattedModuleTemplate += string.Format(_moduleInit, assertSetAs, fakesSetAs);
            }

            if (settings.MethodInit)
            {
                formattedModuleTemplate += _methodInit;
            }

            return formattedModuleTemplate;
        }

        private string InstantiationFormatFor(string type, IUnitTestSettings settings) 
        {
            const string EarlyBoundInstantiationFormat = "New {0}";
            const string LateBoundInstantiationFormat = "CreateObject(\"{0}\")";
            return string.Format(settings.BindingMode == BindingMode.EarlyBinding ? EarlyBoundInstantiationFormat : LateBoundInstantiationFormat, type); 
        }

        private string DeclarationFormatFor(string declarationFormat, string type, IUnitTestSettings settings) 
        {
            return string.Format(declarationFormat, settings.BindingMode == BindingMode.EarlyBinding ? type : "Object");
        }

        private IVBProject GetProject()
        {
            using (var activeProject = _vbe.ActiveVBProject)
            {
                if (!activeProject.IsWrappingNullReference)
                {
                    return activeProject;
                }
            }

            using (var projects = _vbe.VBProjects)
            {
                return projects.Count == 1
                    ? projects[1]
                    : new VBProject(null);
            }
        }

        protected override bool EvaluateCanExecute(object parameter)
        {
            return !GetProject().IsWrappingNullReference && _vbe.HostSupportsUnitTests();
        }

        protected override void OnExecute(object parameter)
        {
            var parameterIsModuleDeclaration = parameter is ProceduralModuleDeclaration || parameter is ClassModuleDeclaration;

            var project = parameter as IVBProject ??
                          (parameterIsModuleDeclaration ? ((Declaration) parameter).Project : GetProject());

            if (project.IsWrappingNullReference)
            {
                return;
            }

            var settings = _configLoader.LoadConfiguration().UserSettings.UnitTestSettings;

            if (settings.BindingMode == BindingMode.EarlyBinding)
            {
                project.EnsureReferenceToAddInLibrary();
            }

            using(var components = project.VBComponents)
            {
                using (var component = components.Add(ComponentType.StandardModule))
                {
                    using (var module = component.CodeModule)
                    {
                        component.Name = GetNextTestModuleName(project);

                        var hasOptionExplicit = false;
                        if (module.CountOfLines > 0 && module.CountOfDeclarationLines > 0)
                        {
                            hasOptionExplicit = module.GetLines(1, module.CountOfDeclarationLines)
                                .Contains("Option Explicit");
                        }

                        var options = string.Concat(hasOptionExplicit ? string.Empty : "Option Explicit\r\n",
                            "Option Private Module\r\n\r\n");

                        if (parameterIsModuleDeclaration)
                        {
                            var moduleCodeBuilder = new StringBuilder();
                            var declarationsToStub = GetDeclarationsToStub((Declaration) parameter);

                            foreach (var declaration in declarationsToStub)
                            {
                                var name = string.Empty;

                                switch (declaration.DeclarationType)
                                {
                                    case DeclarationType.Procedure:
                                    case DeclarationType.Function:
                                        name = declaration.IdentifierName;
                                        break;
                                    case DeclarationType.PropertyGet:
                                        name = $"Get{declaration.IdentifierName}";
                                        break;
                                    case DeclarationType.PropertyLet:
                                        name = $"Let{declaration.IdentifierName}";
                                        break;
                                    case DeclarationType.PropertySet:
                                        name = $"Set{declaration.IdentifierName}";
                                        break;
                                }

                                var stub = AddTestMethodCommand.TestMethodTemplate.Replace(
                                    AddTestMethodCommand.NamePlaceholder, $"{name}_Test");
                                moduleCodeBuilder.AppendLine(stub);
                            }

                            module.AddFromString(options + GetTestModule(settings) + moduleCodeBuilder);
                        }
                        else
                        {
                            var defaultTestMethod = settings.DefaultTestStubInNewModule
                                ? AddTestMethodCommand.TestMethodTemplate.Replace(AddTestMethodCommand.NamePlaceholder,
                                    "TestMethod1")
                                : string.Empty;

                            module.AddFromString(options + GetTestModule(settings) + defaultTestMethod);
                        }
                    }

                    component.Activate();
                }
            }
            _state.OnParseRequested(this);
        }

        private string GetNextTestModuleName(IVBProject project)
        {
            var names = project.ComponentNames();
            var index = names.Count(n => n.StartsWith(TestModuleBaseName)) + 1;

            return string.Concat(TestModuleBaseName, index);
        }

        private IEnumerable<Declaration> GetDeclarationsToStub(Declaration parentDeclaration)
        {
            return _state.DeclarationFinder.Members(parentDeclaration)
                .Where(d => Equals(d.ParentDeclaration, parentDeclaration) && d.Accessibility == Accessibility.Public &&
                            (d.DeclarationType == DeclarationType.Procedure || d.DeclarationType == DeclarationType.Function || d.DeclarationType.HasFlag(DeclarationType.Property)))
                .OrderBy(d => d.Context.Start.TokenIndex);
        }
    }
}