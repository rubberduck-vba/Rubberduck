using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;

namespace Rubberduck.Inspections
{
    public class ParameterCanBeByValInspection : IInspection
    {
        public ParameterCanBeByValInspection()
        {
            Severity = CodeInspectionSeverity.Warning;
        }

        public string Name { get { return RubberduckUI.ParameterCanBeByVal_; } }
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
                && !IsUsedAsByRefParam(parseResult.Declarations, declaration)
                && !declaration.References.Any(reference => reference.IsAssignment))
                .Select(issue => new ParameterCanBeByValInspectionResult(string.Format(Name, issue.IdentifierName), Severity, ((dynamic)issue.Context).ambiguousIdentifier(), issue.QualifiedName));

            return issues;
        }

        private bool IsUsedAsByRefParam(Declarations declarations, Declaration parameter)
        {
            // todo: enable tracking parameter references 
            // by linking Parameter declarations to their parent Procedure/Function/Property member.
            return false;
        }
    }
}