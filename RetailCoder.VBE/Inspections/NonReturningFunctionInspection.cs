using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Common;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections
{
    public sealed class NonReturningFunctionInspection : InspectionBase
    {
        public NonReturningFunctionInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Error)
        {
        }

        public override string Meta { get { return InspectionsUI.NonReturningFunctionInspectionMeta; }}
        public override string Description { get { return InspectionsUI.NonReturningFunctionInspectionName; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }

        private static readonly DeclarationType[] ReturningMemberTypes =
        {
            DeclarationType.Function,
            DeclarationType.PropertyGet
        };

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
        {
            var declarations = UserDeclarations.ToList();

            var interfaceMembers = declarations.FindInterfaceMembers();
            var interfaceImplementationMembers = declarations.FindInterfaceImplementationMembers();

            var functions = declarations
                .Where(declaration => ReturningMemberTypes.Contains(declaration.DeclarationType)
                    && !interfaceMembers.Contains(declaration)).ToList();

            var issues = functions
                .Where(declaration => !declaration.References.Any(reference => reference.IsAssignment && reference.ParentScoping.Equals(declaration)))
                .ToList();
                
            return issues.Select(issue => new NonReturningFunctionInspectionResult(this, issue, interfaceImplementationMembers.Select(m => m.Scope).Contains(issue.Scope)));
        }
    }
}