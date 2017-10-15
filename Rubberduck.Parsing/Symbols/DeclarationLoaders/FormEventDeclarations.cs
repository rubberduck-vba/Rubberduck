using System.Collections.Generic;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Symbols.DeclarationLoaders
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
            var finder = state.DeclarationFinder;

            var msForms = finder.FindProject("MSForms");
            if (msForms == null)
            {
                //The corresponding COM reference has not been loaded.
                return null;
            }

            return finder.FindClassModule("FormEvents", msForms, true);
        }


        private IReadOnlyList<Declaration> AddHiddenMSFormDeclarations(Declaration formsClassModule)
        {

            var userFormActivateEvent = UserFormActivateEvent(formsClassModule);
            var userFormDeactivateEvent = UserFormDeactivateEvent(formsClassModule);
            var userFormInitializeEvent = UserFormInitializeEvent(formsClassModule);
            var userFormQueryCloseEvent = UserFormQueryCloseEvent(formsClassModule);
            var userFormQueryCloseEventCancelParameter = UserFormQueryCloseEventCancelParameter(userFormQueryCloseEvent);
            var userFormQueryCloseEventCloseModeParameter = UserFormQueryCloseEventCloseModeParameter(userFormQueryCloseEvent);
            var userFormResizeEvent = UserFormResizeEvent(formsClassModule);
            var userFormTerminateEvent = UserFormTerminateEvent(formsClassModule);

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

        private static Declaration UserFormActivateEvent(Declaration formsClassModule)
        {
            return new Declaration(
                new QualifiedMemberName(formsClassModule.QualifiedName.QualifiedModuleName, "Activate"),
                formsClassModule,
                formsClassModule.Scope,
                string.Empty,
                string.Empty,
                false,
                false,
                Accessibility.Global,
                DeclarationType.Event,
                false,
                null);
        }

        private static Declaration UserFormDeactivateEvent(Declaration formsClassModule)
        {
            return new Declaration(
                new QualifiedMemberName(formsClassModule.QualifiedName.QualifiedModuleName, "Deactivate"),
                formsClassModule,
                formsClassModule.Scope,
                string.Empty,
                string.Empty,
                false,
                false,
                Accessibility.Global,
                DeclarationType.Event,
                false,
                null);
        }

        private static Declaration UserFormInitializeEvent(Declaration formsClassModule)
        {
            return new Declaration(
                new QualifiedMemberName(formsClassModule.QualifiedName.QualifiedModuleName, "Initialize"),
                formsClassModule,
                formsClassModule.Scope,
                string.Empty,
                string.Empty,
                false,
                false,
                Accessibility.Global,
                DeclarationType.Event,
                false,
                null);
        }

        private static Declaration UserFormQueryCloseEvent(Declaration formsClassModule)
        {
            return new Declaration(
                new QualifiedMemberName(formsClassModule.QualifiedName.QualifiedModuleName, "QueryClose"),
                formsClassModule,
                formsClassModule.Scope,
                string.Empty,
                string.Empty,
                false,
                false,
                Accessibility.Global,
                DeclarationType.Event,
                false,
                null);
        }

        private static ParameterDeclaration UserFormQueryCloseEventCancelParameter(Declaration userFormQueryCloseEvent)
        {
            return new ParameterDeclaration(
                new QualifiedMemberName(userFormQueryCloseEvent.QualifiedName.QualifiedModuleName, "Cancel"),
                userFormQueryCloseEvent,
                null,
                new Selection(),
                "Integer",
                null,
                string.Empty,
                false,
                true,
                false,
                false,
                true);
        }

        private static ParameterDeclaration UserFormQueryCloseEventCloseModeParameter(Declaration userFormQueryCloseEvent)
        {
            return new ParameterDeclaration(
                new QualifiedMemberName(userFormQueryCloseEvent.QualifiedName.QualifiedModuleName, "CloseMode"),
                userFormQueryCloseEvent,
                null,
                new Selection(),
                "Integer",
                null,
                string.Empty,
                false,
                true,
                false,
                false,
                true);
        }

        private static Declaration UserFormResizeEvent(Declaration formsClassModule)
        {
            return new Declaration(
                new QualifiedMemberName(formsClassModule.QualifiedName.QualifiedModuleName, "Resize"),
                formsClassModule,
                formsClassModule.Scope,
                string.Empty,
                string.Empty,
                false,
                false,
                Accessibility.Global,
                DeclarationType.Event,
                false,
                null);
        }

        private static Declaration UserFormTerminateEvent(Declaration formsClassModule)
        {
            return new Declaration(
                new QualifiedMemberName(formsClassModule.QualifiedName.QualifiedModuleName, "Terminate"),
                formsClassModule,
                formsClassModule.Scope,
                string.Empty,
                string.Empty,
                false,
                false,
                Accessibility.Global,
                DeclarationType.Event,
                false,
                null);
        }

    }
}