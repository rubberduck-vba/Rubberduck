using System.Collections.Generic;
using System.Linq;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.Inspections
{
    public sealed class ProcedureNotUsedInspection : InspectionBase
    {
        private readonly IMessageBox _messageBox;

        public ProcedureNotUsedInspection(RubberduckParserState state, IMessageBox messageBox)
            : base(state)
        {
            _messageBox = messageBox;
        }

        public override string Meta { get { return InspectionsUI.ProcedureNotUsedInspectionMeta; } }
        public override string Description { get { return InspectionsUI.ProcedureNotUsedInspectionName; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }

        private static readonly string[] DocumentEventHandlerPrefixes =
        {
            "Chart_",
            "Worksheet_",
            "Workbook_",
            "Document_",
            "Application_",
            "Session_"
        };

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
        {
            var declarations = UserDeclarations.ToList();

            var classes = State.DeclarationFinder.UserDeclarations(DeclarationType.ClassModule).ToList(); // declarations.Where(item => item.DeclarationType == DeclarationType.ClassModule).ToList();
            var modules = State.DeclarationFinder.UserDeclarations(DeclarationType.ProceduralModule).ToList(); // declarations.Where(item => item.DeclarationType == DeclarationType.ProceduralModule).ToList();

            var handlers = State.DeclarationFinder.UserDeclarations(DeclarationType.Control)
                .SelectMany(control => declarations.FindEventHandlers(control)).ToList();

            var builtInHandlers = State.DeclarationFinder.FindEventHandlers();
            handlers.AddRange(builtInHandlers);

            var withEventFields = State.DeclarationFinder.UserDeclarations(DeclarationType.Variable).Where(item => item.IsWithEvents).ToList();
            var withHanders = withEventFields
                .SelectMany(field => State.DeclarationFinder.FindHandlersForWithEventsField(field))
                .ToList();

            handlers.AddRange(withHanders);

            var forms = State.DeclarationFinder.UserDeclarations(DeclarationType.ClassModule)
                .Where(item => item.QualifiedName.QualifiedModuleName.ComponentType == ComponentType.UserForm)
                .ToList();

            if (forms.Any())
            {
                handlers.AddRange(forms.SelectMany(form => State.FindFormEventHandlers(form)));
            }

            var interfaceMembers = State.DeclarationFinder.FindAllInterfaceMembers().ToList();
            var implementingMembers = State.DeclarationFinder.FindAllInterfaceImplementingMembers().ToList();

            var items = declarations
                .Where(item => !IsIgnoredDeclaration(item, interfaceMembers, implementingMembers, handlers, classes, modules)).ToList();
            var issues = items.Select(issue => new IdentifierNotUsedInspectionResult(this, issue, issue.Context, issue.QualifiedName.QualifiedModuleName));

            issues = DocumentEventHandlerPrefixes
                .Aggregate(issues, (current, item) => current.Where(issue => !issue.Description.Contains("'" + item)));

            return issues.ToList();
        }

        private static readonly DeclarationType[] ProcedureTypes =
        {
            DeclarationType.Procedure,
            DeclarationType.Function,
            DeclarationType.LibraryProcedure,
            DeclarationType.LibraryFunction,
            DeclarationType.Event
        };

        private bool IsIgnoredDeclaration(Declaration declaration, IEnumerable<Declaration> interfaceMembers, IEnumerable<Declaration> interfaceImplementingMembers , IEnumerable<Declaration> handlers, IEnumerable<Declaration> classes, IEnumerable<Declaration> modules)
        {
            var enumerable = classes as IList<Declaration> ?? classes.ToList();
            var result = !ProcedureTypes.Contains(declaration.DeclarationType)
                || declaration.References.Any(r => !r.IsAssignment && !r.ParentScoping.Equals(declaration)) // recursive calls don't count
                || handlers.Contains(declaration)
                || IsPublicModuleMember(modules, declaration)
                || IsClassLifeCycleHandler(enumerable, declaration)
                || interfaceMembers.Contains(declaration)
                || interfaceImplementingMembers.Contains(declaration);

            return result;
        }

        /// <remarks>
        /// We cannot determine whether exposed members of standard modules are called or not,
        /// so we assume they are instead of flagging them as "never called".
        /// </remarks>
        private bool IsPublicModuleMember(IEnumerable<Declaration> modules, Declaration procedure)
        {
            if ((procedure.Accessibility != Accessibility.Implicit
                 && procedure.Accessibility != Accessibility.Public))
            {
                return false;
            }

            var parent = modules.Where(item => item.ProjectId == procedure.ProjectId)
                        .SingleOrDefault(item => item.IdentifierName == procedure.ComponentName);

            return parent != null;
        }

        // TODO: Put this into grammar?
        private static readonly string[] ClassLifeCycleHandlers =
        {
            "Class_Initialize",
            "Class_Terminate"
        };

        private bool IsClassLifeCycleHandler(IEnumerable<Declaration> classes, Declaration procedure)
        {
            if (!ClassLifeCycleHandlers.Contains(procedure.IdentifierName))
            {
                return false;
            }

            var parent = classes.Where(item => item.ProjectId == procedure.ProjectId)
                        .SingleOrDefault(item => item.IdentifierName == procedure.ComponentName);

            return parent != null;
        }
    }
}
