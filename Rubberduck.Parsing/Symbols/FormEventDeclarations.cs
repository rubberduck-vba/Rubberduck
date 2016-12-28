using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Symbols
{
    public class FormEventDeclarations : ICustomDeclarationLoader
    {
        private readonly RubberduckParserState _state;

        public FormEventDeclarations(RubberduckParserState state)
        {
            _state = state;
        }

        public IReadOnlyList<Declaration> Load()
        {
            var formsClassModule = FormsClassModuleFromParserState(_state);

            if (formsClassModule == null)
            {
                return new List<Declaration>();
            }

            return AddHiddenMSFormDeclarations(formsClassModule);
        }

        private static Declaration FormsClassModuleFromParserState(RubberduckParserState state)
        {
            var result = state.AllDeclarations.LastOrDefault(declaration =>
                declaration.DeclarationType == DeclarationType.ClassModule
                && declaration.ProjectName == "MSForms"
                && declaration.IdentifierName == "FormEvents");

            return result;
        }

        private IReadOnlyList<Declaration> AddHiddenMSFormDeclarations(Declaration formsClassModule)
        {

            var userFormActivateEvent = CreateDeclaration(formsClassModule, "Activate");
            var userFormDeactivateEvent = CreateDeclaration(formsClassModule, "Deactivate");
            var userFormInitializeEvent = CreateDeclaration(formsClassModule, "Initialize");
            var userFormQueryCloseEvent = CreateDeclaration(formsClassModule, "QueryClose");
            var userFormQueryCloseEventCancelParameter = CreateParameter(userFormQueryCloseEvent, "Cancel", "Integer", true);
            var userFormQueryCloseEventCloseModeParameter = CreateParameter(userFormQueryCloseEvent, "CloseMode", "Integer", true);
            var userFormResizeEvent = CreateDeclaration(formsClassModule, "Resize");
            var userFormTerminateEvent = CreateDeclaration(formsClassModule, "Terminate");

            return new List<Declaration>
            {
                userFormActivateEvent,
                userFormDeactivateEvent,
                userFormInitializeEvent,
                userFormQueryCloseEvent,
                userFormQueryCloseEventCancelParameter,
                userFormQueryCloseEventCloseModeParameter,
                userFormResizeEvent,
                userFormTerminateEvent
            };
        }

        private static Declaration CreateDeclaration(Declaration parent, string name)
        {
            return new Declaration(
                new QualifiedMemberName(parent.QualifiedName.QualifiedModuleName, name),
                parent,
                parent.Scope,
                string.Empty,
                string.Empty,
                false,
                false,
                Accessibility.Global,
                DeclarationType.Event,
                false,
                null);
        }

        private static ParameterDeclaration CreateParameter(Declaration parent, string name, string asType, bool isByRef = false, bool isOptional = false)
        {
            return new ParameterDeclaration(
                new QualifiedMemberName(parent.ParentDeclaration.QualifiedName.QualifiedModuleName, name),
                parent,
                null,
                Selection.Empty,
                asType,
                null,
                string.Empty,
                isOptional,
                isByRef);
        }
    }
}