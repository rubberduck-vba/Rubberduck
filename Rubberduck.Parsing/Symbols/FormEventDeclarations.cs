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
                return state.AllDeclarations.LastOrDefault(declaration => declaration.DeclarationType == DeclarationType.ClassModule
                                                                            && declaration.Scope == "FM20.DLL;MSForms.FormEvents");
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
                        new QualifiedMemberName(
                            new QualifiedModuleName("MSForms", "C:\\WINDOWS\\system32\\FM20.DLL", "FormEvents"), "Activate"),
                        formsClassModule,
                        "FM20.DLL;MSForms.FormEvents",
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
                        new QualifiedMemberName(
                            new QualifiedModuleName("MSForms", "C:\\WINDOWS\\system32\\FM20.DLL", "FormEvents"), "Deactivate"),
                        formsClassModule,
                        "FM20.DLL;MSForms.FormEvents",
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
                        new QualifiedMemberName(
                            new QualifiedModuleName("MSForms", "C:\\WINDOWS\\system32\\FM20.DLL", "FormEvents"), "Initialize"),
                        formsClassModule,
                        "FM20.DLL;MSForms.FormEvents",
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
                        new QualifiedMemberName(
                            new QualifiedModuleName("MSForms", "C:\\WINDOWS\\system32\\FM20.DLL", "FormEvents"), "QueryClose"),
                        formsClassModule,
                        "FM20.DLL;MSForms.FormEvents",
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
                        new QualifiedMemberName(
                            new QualifiedModuleName("MSForms", "C:\\WINDOWS\\system32\\FM20.DLL", "FormEvents"), "Cancel"),
                        userFormQueryCloseEvent,
                        null,
                        new Selection(),
                        "Integer",
                        null,
                        string.Empty,
                        false,
                        true);
                }

                private static ParameterDeclaration UserFormQueryCloseEventCloseModeParameter(Declaration userFormQueryCloseEvent)
                {
                    return new ParameterDeclaration(
                        new QualifiedMemberName(
                            new QualifiedModuleName("MSForms", "C:\\WINDOWS\\system32\\FM20.DLL", "FormEvents"), "CloseMode"),
                        userFormQueryCloseEvent,
                        null,
                        new Selection(),
                        "Integer",
                        null,
                        string.Empty,
                        false,
                        true);
                }

                private static Declaration UserFormResizeEvent(Declaration formsClassModule)
                {
                    return new Declaration(
                        new QualifiedMemberName(
                            new QualifiedModuleName("MSForms", "C:\\WINDOWS\\system32\\FM20.DLL", "FormEvents"), "Resize"),
                        formsClassModule,
                        "FM20.DLL;MSForms.FormEvents",
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
                        new QualifiedMemberName(
                            new QualifiedModuleName("MSForms", "C:\\WINDOWS\\system32\\FM20.DLL", "FormEvents"), "Terminate"),
                        formsClassModule,
                        "FM20.DLL;MSForms.FormEvents",
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