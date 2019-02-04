using System;
using System.Collections.Generic;
using System.Linq;
using NLog;
using Rubberduck.Interaction;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources.UnitTesting;
using Rubberduck.SettingsProvider;
using Rubberduck.SmartIndenter;
using Rubberduck.UnitTesting.Settings;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UnitTesting.UnitTesting
{
    public partial class TestCodeGenerator : ITestCodeGenerator
    {
        protected static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        private readonly bool _isAccess;
        private readonly IConfigProvider<UnitTestSettings> _settings;
        private readonly RubberduckParserState _state;
        private readonly IMessageBox _messageBox;
        private readonly IIndenter _indenter;
        private readonly IVBEInteraction _interaction;

        public TestCodeGenerator(IVBE vbe, RubberduckParserState state, IMessageBox messageBox, IVBEInteraction interaction, IConfigProvider<UnitTestSettings> settings, IIndenter indenter)
        {
            _isAccess = "AccessApp".Equals(vbe?.HostApplication()?.GetType().Name);
            _state = state;
            _messageBox = messageBox;
            _interaction = interaction;
            _settings = settings;           
            _indenter = indenter;
        }

        public void AddTestModuleToProject(IVBProject project)
        {
            AddTestModule(project, null);
        }

        public void AddTestModuleToProject(IVBProject project, Declaration stubSource)
        {
            AddTestModule(project, stubSource);
        }

        private void AddTestModule(IVBProject project, Declaration stubSource)
        {
            if (project == null || project.IsWrappingNullReference)
            {
                return;
            }

            var settings = _settings.Create();

            if (settings.BindingMode == BindingMode.EarlyBinding)
            {
                // FIXME: Push the actual adding of TestModules into UnitTesting, which sidesteps VBEInteraction being inaccessble here
                _interaction.EnsureProjectReferencesUnitTesting(project);
            }

            try
            {
                using (var components = project.VBComponents)
                using (var component = components.Add(ComponentType.StandardModule))
                using (var module = component.CodeModule)
                {
                    component.Name = GetNextTestModuleName(project);

                    // Test modules always have appropriate options so remove any pre-generated code.
                    if (module.CountOfLines > 0)
                    {
                        module.DeleteLines(1, module.CountOfLines);
                    }

                    if (stubSource != null)
                    {
                        var code = GetNewTestModuleCode(component, GetDeclarationsToStub(stubSource).ToList());
                        module.AddFromString(code);
                    }
                    else
                    {
                        module.AddFromString(GetNewTestModuleCode(component));
                    }

                    component.Activate();
                }
            }
            catch (Exception ex)
            {
                _messageBox.Message(TestExplorer.Command_AddTestModule_Error);
                Logger.Warn("Unable to add test module. An exception was thrown.");
                Logger.Warn(ex);
            }
        }

        private string GetNewTestModuleCode(IVBComponent component)
        {
            var settings = _settings.Create();
            var baseModule = (GetBaseTestModule(settings) + Environment.NewLine);

            if (settings.DefaultTestStubInNewModule)
            {
                baseModule += Environment.NewLine + GetNewTestMethodCode(component);
            }

            return string.Join(Environment.NewLine, _indenter.Indent(baseModule));
        }

        private string GetNewTestModuleCode(IVBComponent component, List<Declaration> stubs)
        {
            if (stubs is null || !stubs.Any())
            {
                return GetNewTestModuleCode(component);
            }

            var module = string.Join(Environment.NewLine + Environment.NewLine, new [] { GetBaseTestModule() }.Concat(stubs.Select(GetNewTestStubMethod)));

            return string.Join(Environment.NewLine, _indenter.Indent(module));
        }

        private string GetBaseTestModule(UnitTestSettings settings = null)
        {
            if (settings is null)
            {
                settings = _settings.Create();
            }

            string declaration;
            string initialization;
            var asserts = settings.AssertMode == AssertMode.PermissiveAssert ? "PermissiveAssertClass" : "AssertClass";

            switch (settings.BindingMode)
            {
                case BindingMode.EarlyBinding:
                    declaration = string.Format(EarlyBindingDeclarations, asserts);
                    initialization = string.Format(EarlyBindingInitialization, asserts);
                    break;
                case BindingMode.LateBinding:
                    declaration = LateBindingDeclarations;
                    initialization = string.Format(LateBindingInitialization, asserts);
                    break;
                case BindingMode.DualBinding:
                    declaration = string.Format(DualBindingDeclarations, asserts);
                    initialization = string.Format(DualBindingInitialization, asserts);
                    break;
                default:
                    throw new ArgumentOutOfRangeException();
            }

            return string.Format(TestModuleTemplate,
                _isAccess ? AccessCompareOption : string.Empty,
                declaration,
                initialization);
        }

        public string GetNewTestMethodCode(IVBComponent component)
        {
            return string.Join(Environment.NewLine,
                _indenter.Indent(string.Format(TestMethodTemplate, GetNextTestMethodName(component))));
        }

        public string GetNewTestMethodCodeErrorExpected(IVBComponent component)
        {
            return string.Join(Environment.NewLine,
                _indenter.Indent(string.Format(TestMethodExpectedErrorTemplate, GetNextTestMethodName(component))));
        }

        private string GetNewTestStubMethod(Declaration procedure)
        {
            var name = string.Empty;

            switch (procedure.DeclarationType)
            {
                case DeclarationType.Procedure:
                case DeclarationType.Function:
                    name = $"{procedure.IdentifierName}{TestMethodBaseName}";
                    break;
                case DeclarationType.PropertyGet:
                    name = $"Get{procedure.IdentifierName}{TestMethodBaseName}";
                    break;
                case DeclarationType.PropertyLet:
                    name = $"Let{procedure.IdentifierName}{TestMethodBaseName}";
                    break;
                case DeclarationType.PropertySet:
                    name = $"Set{procedure.IdentifierName}{TestMethodBaseName}";
                    break;
            }

            return string.Format(TestMethodTemplate, name);
        }

        private string GetNextTestModuleName(IVBProject project)
        {
            var names = new HashSet<string>(project.ComponentNames().Where(module => module.StartsWith(TestModuleBaseName)));

            var index = 1;
            while (names.Contains($"{TestModuleBaseName}{index}"))
            {
                index++;
            }

            return $"{TestModuleBaseName}{index}";
        }

        private string GetNextTestMethodName(IVBComponent component)
        {
            var names = new HashSet<string>(_state.DeclarationFinder.Members(component.QualifiedModuleName)
                .Select(test => test.IdentifierName).Where(decl => decl.StartsWith(TestMethodBaseName)));

            var index = 1;
            while (names.Contains($"{TestMethodBaseName}{index}"))
            {
                index++;
            }

            return $"{TestMethodBaseName}{index}";
        }

        private IEnumerable<Declaration> GetDeclarationsToStub(Declaration parentDeclaration)
        {
            return _state.DeclarationFinder.Members(parentDeclaration)
                .Where(d =>
                            Equals(d.ParentDeclaration, parentDeclaration) && d.Accessibility == Accessibility.Public &&
                            (d.DeclarationType == DeclarationType.Procedure || d.DeclarationType == DeclarationType.Function ||
                             d.DeclarationType.HasFlag(DeclarationType.Property)))
                .OrderBy(d => d.Context.Start.TokenIndex);
        }
    }
}
