using System.Collections.Generic;
using NLog;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Symbols
{
    public class FormEventDeclarations : ICustomDeclarationLoader
    {
        private readonly RubberduckParserState _state;
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

        public FormEventDeclarations(RubberduckParserState state)
        {
            _state = state;
        }

        public IReadOnlyList<Declaration> Load()
        {
            Declaration formsClassModule = null;
            foreach (var declaration in _state.AllDeclarations)
            {
                if (declaration.DeclarationType == DeclarationType.ClassModule &&
                    declaration.Scope == "FM20.DLL;MSForms.FormEvents")
                {
                    formsClassModule = declaration;
                }
            }

            if (formsClassModule == null)
            {
                return new List<Declaration>();
            }

            return AddHiddenMSFormDeclarations(formsClassModule);
        }

        private IReadOnlyList<Declaration> AddHiddenMSFormDeclarations(Declaration parentModule)
        {
            var userFormActivateEvent = new Declaration(
                new QualifiedMemberName(
                    new QualifiedModuleName("MSForms", "C:\\WINDOWS\\system32\\FM20.DLL", "FormEvents"), "Activate"),
                parentModule,
                "FM20.DLL;MSForms.FormEvents",
                string.Empty,
                string.Empty,
                false,
                false,
                Accessibility.Global,
                DeclarationType.Event,
                false,
                null);

            var userFormDeactivateEvent = new Declaration(
                new QualifiedMemberName(
                    new QualifiedModuleName("MSForms", "C:\\WINDOWS\\system32\\FM20.DLL", "FormEvents"), "Deactivate"),
                parentModule,
                "FM20.DLL;MSForms.FormEvents",
                string.Empty,
                string.Empty,
                false,
                false,
                Accessibility.Global,
                DeclarationType.Event,
                false,
                null);

            var userFormInitializeEvent = new Declaration(
                new QualifiedMemberName(
                    new QualifiedModuleName("MSForms", "C:\\WINDOWS\\system32\\FM20.DLL", "FormEvents"), "Initialize"),
                parentModule,
                "FM20.DLL;MSForms.FormEvents",
                string.Empty,
                string.Empty,
                false,
                false,
                Accessibility.Global,
                DeclarationType.Event,
                false,
                null);

            var userFormQueryCloseEvent = new Declaration(
                new QualifiedMemberName(
                    new QualifiedModuleName("MSForms", "C:\\WINDOWS\\system32\\FM20.DLL", "FormEvents"), "QueryClose"),
                parentModule,
                "FM20.DLL;MSForms.FormEvents",
                string.Empty,
                string.Empty,
                false,
                false,
                Accessibility.Global,
                DeclarationType.Event,
                false,
                null);

            var userFormQueryCloseEventCancelParameter = new ParameterDeclaration(
                new QualifiedMemberName(
                    new QualifiedModuleName("MSForms", "C:\\WINDOWS\\system32\\FM20.DLL", "FormEvents"), "Cancel"),
                userFormQueryCloseEvent,
                null,
                new Selection(),
                "Integer",
                null,
                string.Empty,
                false,
                false);

            var userFormQueryCloseEventCloseModeParameter = new ParameterDeclaration(
                new QualifiedMemberName(
                    new QualifiedModuleName("MSForms", "C:\\WINDOWS\\system32\\FM20.DLL", "FormEvents"), "CloseMode"),
                userFormQueryCloseEvent,
                null,
                new Selection(),
                "Integer",
                null,
                string.Empty,
                false,
                false);

            var userFormResizeEvent = new Declaration(
                new QualifiedMemberName(
                    new QualifiedModuleName("MSForms", "C:\\WINDOWS\\system32\\FM20.DLL", "FormEvents"), "Resize"),
                parentModule,
                "FM20.DLL;MSForms.FormEvents",
                string.Empty,
                string.Empty,
                false,
                false,
                Accessibility.Global,
                DeclarationType.Event,
                false,
                null);

            var userFormTerminateEvent = new Declaration(
                new QualifiedMemberName(
                    new QualifiedModuleName("MSForms", "C:\\WINDOWS\\system32\\FM20.DLL", "FormEvents"), "Terminate"),
                parentModule,
                "FM20.DLL;MSForms.FormEvents",
                string.Empty,
                string.Empty,
                false,
                false,
                Accessibility.Global,
                DeclarationType.Event,
                false,
                null);

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
    }
}