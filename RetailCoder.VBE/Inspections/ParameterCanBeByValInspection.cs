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
            Severity = CodeInspectionSeverity.Warning;
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
                .Concat(parseResult.Declarations.FindInterfaceImplementationMembers())
                .ToList();

            var formEventHandlerScopes = parseResult.Declarations.FindFormEventHandlers()
                .Select(handler => handler.Scope);

            var eventScopes = parseResult.Declarations.Items.Where(item => 
                !item.IsBuiltIn && item.DeclarationType == DeclarationType.Event)
                .Select(e => e.Scope);

            var declareScopes = parseResult.Declarations.Items.Where(item => 
                    item.DeclarationType == DeclarationType.LibraryFunction 
                    || item.DeclarationType == DeclarationType.LibraryProcedure)
                .Select(e => e.Scope);

            var ignoredScopes = formEventHandlerScopes.Concat(eventScopes).Concat(declareScopes);

            var issues = parseResult.Declarations.Items.Where(declaration =>
                !ignoredScopes.Contains(declaration.ParentScope)
                && declaration.DeclarationType == DeclarationType.Parameter
                && !interfaceMembers.Select(m => m.Scope).Contains(declaration.ParentScope)
                && PrimitiveTypes.Contains(declaration.AsTypeName)
                && ((VBAParser.ArgContext) declaration.Context).BYVAL() == null
                && !declaration.References.Any(reference => reference.IsAssignment))
                .Select(issue => new ParameterCanBeByValInspectionResult(string.Format(Name, issue.IdentifierName), Severity, issue.Context, issue.QualifiedName));

            return issues;
        }


    }
}