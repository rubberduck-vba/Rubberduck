using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Concrete
{
    [SuppressMessage("ReSharper", "LoopCanBeConvertedToQuery")]
    public sealed class UnassignedVariableUsageInspection : InspectionBase
    {
        public UnassignedVariableUsageInspection(RubberduckParserState state)
            : base(state) { }

        //See https://github.com/rubberduck-vba/Rubberduck/issues/2010 for why these are being excluded.
        private static readonly List<string> IgnoredFunctions = new List<string>
        {
            "VBE7.DLL;VBA.Strings.Len",
            "VBE7.DLL;VBA.Strings.LenB",
            "VBA6.DLL;VBA.Strings.Len",
            "VBA6.DLL;VBA.Strings.LenB"
        };

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var declarations = State.DeclarationFinder.UserDeclarations(DeclarationType.Variable)
                .Where(declaration => !declaration.IsArray &&
                    State.DeclarationFinder.MatchName(declaration.AsTypeName)
                        .All(d => d.DeclarationType != DeclarationType.UserDefinedType)
                    && !declaration.IsSelfAssigned
                    && !declaration.References.Any(reference => reference.IsAssignment));

            var excludedDeclarations = BuiltInDeclarations.Where(decl => IgnoredFunctions.Contains(decl.QualifiedName.ToString())).ToList();

            return declarations
                .Where(d => d.References.Any() && !excludedDeclarations.Any(excl => DeclarationReferencesContainsReference(excl, d)))
                .SelectMany(d => d.References.Where(r => !IsAssignedByRefArgument(r.ParentScoping, r)))
                .Distinct()
                .Where(r => !r.IsIgnoringInspectionResultFor(AnnotationName))
                .Where(r => !r.Context.TryGetAncestor<VBAParser.RedimStmtContext>(out _) && !IsArraySubscriptAssignment(r))
                .Select(r => new IdentifierReferenceInspectionResult(this,
                    string.Format(InspectionResults.UnassignedVariableUsageInspection, r.IdentifierName),
                    State,
                    r)).ToList();
        }

        private bool IsAssignedByRefArgument(Declaration enclosingProcedure, IdentifierReference reference)
        {
            var argExpression = reference.Context.GetAncestor<VBAParser.ArgumentExpressionContext>();
            var parameter = State.DeclarationFinder.FindParameterOfNonDefaultMemberFromSimpleArgumentNotPassedByValExplicitly(argExpression, enclosingProcedure);

            // note: not recursive, by design.
            return parameter != null
                   && (parameter.IsImplicitByRef || parameter.IsByRef)
                   && parameter.References.Any(r => r.IsAssignment);
        }

        private static bool IsArraySubscriptAssignment(IdentifierReference reference)
        {
            var isLetAssignment = reference.Context.TryGetAncestor<VBAParser.LetStmtContext>(out var letStmt);
            var isSetAssignment = reference.Context.TryGetAncestor<VBAParser.SetStmtContext>(out var setStmt);

            return isLetAssignment && letStmt.lExpression() is VBAParser.IndexExprContext ||
                   isSetAssignment && setStmt.lExpression() is VBAParser.IndexExprContext;
        }

        private static bool DeclarationReferencesContainsReference(Declaration parentDeclaration, Declaration target)
        {
            foreach (var targetReference in target.References)
            {
                foreach (var reference in parentDeclaration.References)
                {
                    var context = (ParserRuleContext)reference.Context.Parent;
                    if (context.GetSelection().Contains(targetReference.Selection))
                    {
                        return true;
                    }
                }
            }

            return false;
        }
    }
}
