using System.Collections.Generic;
using Rubberduck.Parsing;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections
{
    public class ParameterCanBeByValInspection : IInspection
    {
        public ParameterCanBeByValInspection()
        {
            Severity = CodeInspectionSeverity.Suggestion;
        }

        public string Name { get { return InspectionNames.ParameterCanBeByVal_; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        private static string[] PrimitiveTypes =
        {
            Tokens.Boolean,
            Tokens.Byte,
            Tokens.Date,
            Tokens.Decimal,
            Tokens.Double,
            Tokens.Long,
            Tokens.LongLong,
            Tokens.LongPtr,
            Tokens.Integer,
            Tokens.Single,
            Tokens.String,
            Tokens.StrPtr
        };

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(VBProjectParseResult parseResult)
        {
            var interfaceMembers = parseResult.Declarations.FindInterfaceMembers()
                .Concat(parseResult.Declarations.FindInterfaceImplementationMembers());

            var issues = parseResult.Declarations.Items.Where(declaration =>
                declaration.DeclarationType == DeclarationType.Parameter
                && !interfaceMembers.Select(m => m.Scope).Contains(declaration.ParentScope)
                && PrimitiveTypes.Contains(declaration.AsTypeName)
                && ((VBAParser.ArgContext) declaration.Context).BYVAL() == null
                && !declaration.References.Any(reference => reference.IsAssignment))
                .Select(issue => new ParameterCanBeByValInspectionResult(string.Format(Name, issue.IdentifierName), Severity, issue.Context, issue.QualifiedName));

            return issues;
        }
    }
}