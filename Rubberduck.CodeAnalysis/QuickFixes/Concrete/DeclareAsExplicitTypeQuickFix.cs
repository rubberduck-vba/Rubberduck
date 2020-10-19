using Antlr4.Runtime;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.Inspections.Concrete.UnreachableCaseEvaluation;
using Rubberduck.CodeAnalysis.QuickFixes.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.CodeAnalysis.QuickFixes.Concrete
{
    //Expressions are not evaluated.  So, if some assignments are made using expressions, the number 
    //of assignments will not match the inferred AsTypeNames count. In these cases,
    //the QuickFix uses Variant since the analysis is incomplete.  
    internal sealed class DeclareAsExplicitTypeQuickFix : QuickFixBase
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IParseTreeValueFactory _parseTreeValueFactory;

        private QualifiedModuleName _qmn;

        public DeclareAsExplicitTypeQuickFix(IDeclarationFinderProvider declarationFinderProvider, IParseTreeValueFactory parseTreeValueFactory)
            : base(typeof(VariableTypeNotDeclaredInspection))
        {
            _declarationFinderProvider = declarationFinderProvider;
            _parseTreeValueFactory = parseTreeValueFactory;
        }

        public override bool CanFixMultiple => true;
        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => false;
        public override bool CanFixAll => true;

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            _qmn = result.Target.QualifiedModuleName;

            var (identifierContext, asTypeName) = InferTypeForInspectionResult(result);

            var rewriter = rewriteSession.CheckOutModuleRewriter(_qmn);
            rewriter.InsertAfter(identifierContext.Stop.TokenIndex, $" {Tokens.As} {asTypeName}");
        }

        //TODO: Introduce resource
        public override string Description(IInspectionResult result) => "Declare with Explicit Type";

        private (ParserRuleContext identifierContext, string asTypeName) InferTypeForInspectionResult(IInspectionResult result)
        {
            string asTypeName = null;

            switch (result.Target.DeclarationType)
            {
                case DeclarationType.Variable:
                    asTypeName = GetAsTypeNameVariable(result);
                    break;
                case DeclarationType.Constant:
                    asTypeName = GetAsTypeNameConstant(result);
                    break;
                case DeclarationType.Parameter:
                    asTypeName = GetAsTypeNameParameter(result);
                    break;
            }

            return (GetIdentifierContext(result), asTypeName ?? Tokens.Variant);
        }

        private string GetAsTypeNameVariable(IInspectionResult result)
        {
            string asTypeName = null;

            if (result.Target.References.Any())
            {
                var (AsTypeNames, NumberOfAssignmentsFound) = InferTypeNamesFromAssignments(result);

                //TODO: Update once a unified Expression engine is available.
                if (NumberOfAssignmentsFound == AsTypeNames.Count())
                {
                    //TODO: test for scenario where undefined type is used as a ByRef param
                    //to initialize...this should push the parameters evaluation out of this References.Any()
                    var parameterTypes = InferTypeNamesFromParameterUsage(result);

                    asTypeName = ChooseAsTypeNameFromCandidates(AsTypeNames, parameterTypes);
                }

                var usesAtLeastOneLetAssignment = result.Target.References.Any(rf => rf.IsAssignment && !rf.IsSetAssignment);

                return asTypeName
                    ?? (!usesAtLeastOneLetAssignment
                        ? Tokens.Object
                        : Tokens.Variant);
            }
            return asTypeName;
        }

        private string GetAsTypeNameParameter(IInspectionResult result)
        {
            string asTypeName = null;

            var parameterAsTypes = InferTypeNamesFromParameterUsage(result);

            var declarationAsTypes = InferTypeNamesFromDeclarationStatementParameter(result);

            if (result.Target.References.Any())
            {
                var (AsTypeNames, NumberOfAssignmentsFound) = InferTypeNamesFromAssignments(result);

                //TODO: Update once a unified Expression engine is available.
                if (NumberOfAssignmentsFound == AsTypeNames.Count())
                {
                    asTypeName = ChooseAsTypeNameFromCandidates(AsTypeNames, parameterAsTypes, declarationAsTypes);
                }
                var usesAtLeastOneLetAssignment = result.Target.References.Any(rf => rf.IsAssignment && !rf.IsSetAssignment);

                return asTypeName
                    ?? (!usesAtLeastOneLetAssignment
                        ? Tokens.Object
                        : Tokens.Variant);
            }

            return ChooseAsTypeNameFromCandidates(parameterAsTypes, declarationAsTypes);
        }

        private string GetAsTypeNameConstant(IInspectionResult result)
        {
            var parameterAsTypes = InferTypeNamesFromParameterUsage(result);

            var declarationAsTypes = InferTypeNamesFromDeclarationStatement(result, result.Context);

            return ChooseAsTypeNameFromCandidates(parameterAsTypes, declarationAsTypes);
        }

        private List<string> InferTypeNamesFromDeclarationStatementParameter(IInspectionResult result)
        {
            if (result.Context is VBAParser.ArgContext argContext)
            {
                var argDefaultValue = argContext.GetChild<VBAParser.ArgDefaultValueContext>();
                return InferTypeNamesFromDeclarationStatement(result, argDefaultValue as ParserRuleContext);
            }
            return new List<string>();
        }

        private List<string> InferTypeNamesFromDeclarationStatement(IInspectionResult result, ParserRuleContext context)
        {
            string lExprTypeName = null;
            string literalExprTypeName = null;

            if (context.TryGetChildContext<VBAParser.LExprContext>(out var lExpr))
            {
                lExprTypeName = ExtractAsTypeNameFromLExpr(new List<VBAParser.LExprContext>() { lExpr })?.FirstOrDefault();
            }

            if (context.TryGetChildContext<VBAParser.LiteralExprContext>(out var litExpr))
            {
                literalExprTypeName = ExtractAsTypeNameFromLiteral(new List<VBAParser.LiteralExprContext>() { litExpr })?.FirstOrDefault();
            }

            return new List<string>()
            {
                lExprTypeName,
                literalExprTypeName
            };
        }

        private (List<string> AsTypeNames, int NumberOfAssignmentsFound) InferTypeNamesFromAssignments(IInspectionResult result)
        {
            var assignmentTypes = new List<string>();

            var letAssignmentCtxts = result.Target.References
                .Where(rf => rf.IsAssignment && !rf.IsSetAssignment)
                .Select(rf => rf.Context.GetAncestor<VBAParser.LetStmtContext>())
                .Cast<ParserRuleContext>();

            var setAssignmentCtxts = result.Target.References
                .Where(rf => rf.IsAssignment && rf.IsSetAssignment)
                .Select(rf => rf.Context.GetAncestor<VBAParser.SetStmtContext>())
                .Cast<ParserRuleContext>();

            var assignmentContextsToEvaluate = letAssignmentCtxts.Concat(setAssignmentCtxts);

            var numberOfAssignmentContexts = assignmentContextsToEvaluate.Count();

            if (assignmentContextsToEvaluate.Any())
            {
                var lExprTypes = ExtractAsTypeNameFromLExpr(assignmentContextsToEvaluate.Where(ac => ac.TryGetChildContext<VBAParser.LExprContext>(out _))
                    .Select(ac => ac.GetDescendent<VBAParser.LExprContext>())) ?? new List<string>();

                var literalExprTypes = ExtractAsTypeNameFromLiteral(assignmentContextsToEvaluate.Where(ac => ac.TryGetChildContext<VBAParser.LiteralExprContext>(out _))
                    .Select(ac => ac?.GetDescendent<VBAParser.LiteralExprContext>())) ?? new List<string>();

                assignmentTypes.AddRange(lExprTypes.Concat(literalExprTypes));
            }
            return (assignmentTypes, numberOfAssignmentContexts);
        }

        private List<string> InferTypeNamesFromParameterUsage(IInspectionResult result)
        {
            var parameterAsTypeNames = new List<string>();

            var argListContexts = result.Target.References
                .Select(rf => rf.Context.GetAncestor<VBAParser.ArgumentListContext>());

            if (!argListContexts?.Any() ?? true)
            {
                return parameterAsTypeNames;
            }

            foreach (var argListContext in argListContexts)
            {
                (int? ZeroBasedParameterPosition, string ParameterName) = GetParameterLocatorAttributes(argListContext, result.Target.IdentifierName);

                if (!ZeroBasedParameterPosition.HasValue && string.IsNullOrEmpty(ParameterName))
                {
                    parameterAsTypeNames = new List<string>();
                    break;
                }
                var contextOfInterest = argListContext.GetAncestor<VBAParser.IndexExprContext>() as ParserRuleContext
                    ?? argListContext.GetAncestor<VBAParser.CallStmtContext>() as ParserRuleContext;

                var moduleBodyElement = FindProcedureDeclarationFromReferenceContext(contextOfInterest, contextOfInterest.children[0].GetText());

                if (ZeroBasedParameterPosition > moduleBodyElement.Parameters.Count() - 1
                    && moduleBodyElement.Parameters.Last().IsParamArray)
                {
                    parameterAsTypeNames.Add(moduleBodyElement.Parameters.Last().AsTypeName);
                }
                else
                {
                    var parameterAsTypeName = ZeroBasedParameterPosition >= 0
                        ? moduleBodyElement.Parameters.ElementAt(ZeroBasedParameterPosition.Value).AsTypeName
                        : moduleBodyElement.Parameters.Single(p => p.IdentifierName == ParameterName).AsTypeName;

                    parameterAsTypeNames.Add(parameterAsTypeName);
                }
            }

            return parameterAsTypeNames;
        }

        private ModuleBodyElementDeclaration FindProcedureDeclarationFromReferenceContext<T>(T context, string identifier) where T: ParserRuleContext
        {
            var procedure = _declarationFinderProvider.DeclarationFinder
                    .MatchName(identifier)
                    .Where(d => d.DeclarationType.HasFlag(DeclarationType.Function)
                        || d.DeclarationType.HasFlag(DeclarationType.Procedure));

            var referencingProcedureDeclaration = procedure
                .SingleOrDefault(f => f.References.Any(rf => rf.Context.GetAncestor<T>() == context));

            return referencingProcedureDeclaration as ModuleBodyElementDeclaration;
        }

        private static (int? Position, string ArgumentName) GetParameterLocatorAttributes(VBAParser.ArgumentListContext argListContext, string identifier)
        {
            (int? Position, string ArgumentName) result = (null, null);

            var argumentContexts = argListContext.GetDescendents<VBAParser.ArgumentContext>();
            for (var idx = 0; idx < argumentContexts.Count(); idx++)
            {
                var argumentContext = argumentContexts.ElementAt(idx);

                if (!argumentContext.GetText().Contains(identifier))
                {
                    continue;
                }

                if (argumentContext.TryGetChildContext<VBAParser.PositionalArgumentContext>(out _)
                    && argumentContext.GetText().Equals(identifier))
                {
                    result.Position = idx;
                    break;
                }

                if (argumentContext.TryGetChildContext<VBAParser.NamedArgumentContext>(out var namedArgCtxt)
                    && namedArgCtxt.GetChild<VBAParser.ArgumentExpressionContext>().GetText().Equals(identifier))
                {
                    result.ArgumentName = namedArgCtxt.GetChild<VBAParser.UnrestrictedIdentifierContext>().GetText();
                    break;
                }
            }

            return result;
        }

        private static List<string> _numericTypePrecedence = new List<string>()
        {
            Tokens.Double,
            Tokens.Single,
            Tokens.LongLong,
            Tokens.Long,
            Tokens.Integer,
            Tokens.Byte,
            Tokens.Boolean
        };

        //If every inferred AsTypeName is the same, then these AsTypeNames can be used..otherwise, use Variant
        private static List<string> _allOrNothingTypeNames = new List<string>()
        {
            Tokens.String,
            Tokens.Date,
            Tokens.Currency,
            Tokens.Decimal
        };

        private string ChooseAsTypeNameFromCandidates(params List<string>[] theCandidates)// IEnumerable<string> candidate)
        {
            var candidates = new List<string>();
            foreach (var candidateList in theCandidates)
            {
                candidates.AddRange(candidateList);
            }
            var asTypeNames = candidates.Where(t => t != null).ToList();

            if (!asTypeNames.Any())
            {
                return null;
            }

            var asTypes = asTypeNames.ToLookup(asTypeName => asTypeName);
            if (asTypes.Count() == 1)
            {
                //All evaluations of the implicit element resolve to the same AsTypeName...the happy path
                //This path will also return AsTypeNames for Objects
                return asTypes.First().Key;
            }

            //Various AsTypeNames inferred...find the most accomodating

            return asTypes.Any(n => _allOrNothingTypeNames.Contains(n.Key))
                || asTypes.Any(n => !_numericTypePrecedence.Contains(n.Key))
                    ? null
                    : _numericTypePrecedence.Find(t => asTypes.Contains(t));
        }

        private List<string> ExtractAsTypeNameFromLExpr(IEnumerable<VBAParser.LExprContext> lExprContexts)
        {
            if (!lExprContexts.Any() || !lExprContexts.Any(ex => ex != null))
            {
                return null;
            }

            bool IsAccessibleExternally(Declaration d) 
                => d.Accessibility != Accessibility.Private || d.Accessibility != Accessibility.Implicit;

            bool IsReferencedByAnAssignmentInModule(Declaration d)
                => d.References.Any(rf => !rf.IsAssignment
                    && d.QualifiedModuleName != _qmn && rf.QualifiedModuleName == _qmn);

            var externalUserDefinedfunctions = _declarationFinderProvider.DeclarationFinder.UserDeclarations(DeclarationType.Function)
                .Where(f => IsAccessibleExternally(f) && IsReferencedByAnAssignmentInModule(f));

            var libFunction = _declarationFinderProvider.DeclarationFinder.AllBuiltInDeclarations
                .Where(f => f.References.Any(rf => !rf.IsAssignment && rf.QualifiedModuleName == _qmn));

            var externalFields = _declarationFinderProvider.DeclarationFinder.UserDeclarations(DeclarationType.Variable)
                .Where(f => IsAccessibleExternally(f) && IsReferencedByAnAssignmentInModule(f));

            var externalConstants = _declarationFinderProvider.DeclarationFinder.UserDeclarations(DeclarationType.Constant)
                .Where(f => IsAccessibleExternally(f) && IsReferencedByAnAssignmentInModule(f));

            var members = _declarationFinderProvider.DeclarationFinder.Members(_qmn);

            var allAssignmentReferenceCandidates = externalUserDefinedfunctions.Concat(libFunction)
                .Concat(externalFields)
                .Concat(externalConstants)
                .Concat(members)
                .SelectMany(d => d.References.Where(rf => !rf.IsAssignment));

            var referenceAndLExprPairs = allAssignmentReferenceCandidates.Select(mrf => (mrf, mrf.Context.GetAncestor<VBAParser.LExprContext>()));
            var theRHSReferences = referenceAndLExprPairs.Where(refAndLExprPair => lExprContexts.Contains(refAndLExprPair.Item2))
                .Select(lr => lr.mrf);

            return theRHSReferences.Select(rf => rf.Declaration.AsTypeName).ToList();
        }

        private List<string> ExtractAsTypeNameFromLiteral(IEnumerable<VBAParser.LiteralExprContext> literalExprContexts)
        {
            if (!literalExprContexts.Any())
            {
                return null;
            }

            var typeOptions = new List<IParseTreeValue>();
            foreach (var litExpr in literalExprContexts.Where(le => le != null))
            {
                typeOptions.Add(_parseTreeValueFactory.Create(litExpr.GetText()));
            }

            return typeOptions.Select(to => to.ValueType).ToList();
        }

        private static ParserRuleContext GetIdentifierContext(IInspectionResult result)
        {
            switch (result.Context)
            {
                case VBAParser.VariableSubStmtContext _:
                case VBAParser.ConstSubStmtContext _:
                    return result.Context.children[0] as ParserRuleContext;
                case VBAParser.ArgContext argContext:
                    return argContext.unrestrictedIdentifier();
                default:
                    return null;
            }
        }
    }
}
