using System.Collections.Generic;
using System.Linq;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Inspections.Inspections.Extensions;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Identifies parameter declarations that are not used.
    /// </summary>
    /// <why>
    /// Declarations that are not used anywhere should probably be removed.
    /// </why>
    /// <remarks>
    /// Not all unused parameters can/should be removed: ignore any inspection results for 
    /// event handler procedures and interface members that Rubberduck isn't recognizing as such.
    /// </remarks>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething(ByVal foo As Long, ByVal bar As Long)
    ///     Debug.Print foo
    /// End Sub
    /// ]]>
    /// </example>
    /// <example>
    /// <![CDATA[
    /// Option Explicit
    /// Public Sub DoSomething(ByVal foo As Long, ByVal bar As Long)
    ///     Debug.Print foo, bar
    /// End Sub
    /// ]]>
    /// </example>
    public sealed class ParameterNotUsedInspection : InspectionBase
    {
        public ParameterNotUsedInspection(RubberduckParserState state)
            : base(state) { }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var interfaceMembers = State.DeclarationFinder.FindAllInterfaceMembers();
            var interfaceImplementationMembers = State.DeclarationFinder.FindAllInterfaceImplementingMembers();

            var handlers = State.DeclarationFinder.FindEventHandlers();

            var parameters = State.DeclarationFinder
                .UserDeclarations(DeclarationType.Parameter)
                .OfType<ParameterDeclaration>()
                .Where(parameter => !parameter.References.Any()
                                    && parameter.ParentDeclaration.DeclarationType != DeclarationType.Event
                                    && parameter.ParentDeclaration.DeclarationType != DeclarationType.LibraryFunction
                                    && parameter.ParentDeclaration.DeclarationType != DeclarationType.LibraryProcedure
                                    && !interfaceMembers.Contains(parameter.ParentDeclaration)
                                    && !handlers.Contains(parameter.ParentDeclaration))
                .ToList();

            var issues = from issue in parameters
                let isInterfaceImplementationMember = interfaceImplementationMembers.Contains(issue.ParentDeclaration)
                select new DeclarationInspectionResult(this, string.Format(InspectionResults.ParameterNotUsedInspection, issue.IdentifierName).Capitalize(), issue);

            return issues;
        }
    }
}
