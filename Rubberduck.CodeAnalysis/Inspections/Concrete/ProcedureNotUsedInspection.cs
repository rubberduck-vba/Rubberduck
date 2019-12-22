using System.Collections.Generic;
using System.Linq;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Inspections.Extensions;
using Rubberduck.Inspections.Results;
using Rubberduck.JunkDrawer.Extensions;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Extensions;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Locates procedures that are never invoked from user code.
    /// </summary>
    /// <why>
    /// Unused procedures are dead code that should probably be removed. Note, a procedure may be effectively "not used" in code, but attached to some
    /// Shape object in the host document: in such cases the inspection result should be ignored. An event handler procedure that isn't being
    /// resolved as such, may also wrongly trigger this inspection.
    /// </why>
    /// <remarks>
    /// Not all unused procedures can/should be removed: ignore any inspection results for 
    /// event handler procedures and interface members that Rubberduck isn't recognizing as such.
    /// </remarks>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething()
    ///     ' macro is attached to a worksheet Shape.
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Option Explicit
    ///
    /// '@Ignore ProcedureNotUsed
    /// Public Sub DoSomething()
    ///     ' macro is attached to a worksheet Shape.
    /// End Sub
    /// ]]>
    /// </example>
    public sealed class ProcedureNotUsedInspection : InspectionBase
    {
        public ProcedureNotUsedInspection(RubberduckParserState state) : base(state) { }

        private static readonly string[] DocumentEventHandlerPrefixes =
        {
            "Chart_",
            "Worksheet_",
            "Workbook_",
            "Document_",
            "Application_",
            "Session_"
        };

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var classes = State.DeclarationFinder.UserDeclarations(DeclarationType.ClassModule)
                .Concat(State.DeclarationFinder.UserDeclarations(DeclarationType.Document))
                .ToList(); 
            var modules = State.DeclarationFinder.UserDeclarations(DeclarationType.ProceduralModule).ToList();

            var handlers = State.DeclarationFinder.FindEventHandlers().ToHashSet();

            var interfaceMembers = State.DeclarationFinder.FindAllInterfaceMembers().ToHashSet();
            var implementingMembers = State.DeclarationFinder.FindAllInterfaceImplementingMembers().ToHashSet();

            var items = State.AllUserDeclarations
                .Where(item => !IsIgnoredDeclaration(item, interfaceMembers, implementingMembers, handlers, classes, modules))
                .ToList();
            var issues = items.Select(issue => new DeclarationInspectionResult(this,
                                                                    string.Format(InspectionResults.IdentifierNotUsedInspection, issue.DeclarationType.ToLocalizedString(), issue.IdentifierName),
                                                                    issue));

            issues = DocumentEventHandlerPrefixes
                .Aggregate(issues, (current, item) => current.Where(issue => !issue.Description.Contains($"'{item}")));

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
