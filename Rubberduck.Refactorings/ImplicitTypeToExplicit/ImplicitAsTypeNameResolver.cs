using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.ImplicitTypeToExplicit
{
    internal class ImplicitAsTypeNameResolver
    {
        private readonly LiteralExprContextToAsTypeNameConverter _literalExprContextEvaluator;
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly ConcatOpContextResolver _concatOpContextResolver;
        private readonly Declaration _target;

        public ImplicitAsTypeNameResolver(IDeclarationFinderProvider declarationFinderProvider,
            IParseTreeValueFactory parseTreeValueFactory, Declaration target)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _literalExprContextEvaluator = new LiteralExprContextToAsTypeNameConverter(parseTreeValueFactory);
            _target = target;

            _concatOpContextResolver = new ConcatOpContextResolver(declarationFinderProvider);
        }

        public List<string> InferAsTypeNames(IReadOnlyCollection<VBAParser.LiteralExprContext> tContexts) 
            => tContexts.Select(lx => _literalExprContextEvaluator.InferAsTypeName(lx)).ToList();

        public List<string> InferAsTypeNames(IReadOnlyCollection<VBAParser.NewExprContext> tContexts) 
            => tContexts.Select(c => c.GetChild<VBAParser.LExprContext>()?.GetText() ?? string.Empty).ToList();

        public List<string> InferAsTypeNames(IReadOnlyCollection<VBAParser.ConcatOpContext> tContexts)
        {
            return _concatOpContextResolver.InferAsTypeNames(tContexts);
        }

        public List<string> InferAsTypeNames(IReadOnlyCollection<VBAParser.LetStmtContext> tContexts)
        {
            var functionAsTypeNames = new List<string>();
            foreach (var procedureLetCtxt in tContexts)
            {
                var procedureLHSContext = procedureLetCtxt.GetChild(0);
                var procedureIdentifier = procedureLHSContext is VBAParser.MemberAccessExprContext
                    || procedureLHSContext is VBAParser.WithMemberAccessExprContext
                        ? procedureLHSContext.GetChild(procedureLHSContext.ChildCount - 1).GetText()
                        : procedureLHSContext.GetChild(0).GetText();

                var valueParamAsTypeNames = AsTypeNameForValueParameters(procedureIdentifier, procedureLetCtxt);

                if (valueParamAsTypeNames.Any())
                {
                    functionAsTypeNames.AddRange(valueParamAsTypeNames);
                }

                var functionReturnAssignmentTypes = AsTypeNameForFunctionTypes(procedureIdentifier, procedureLetCtxt);

                if (functionReturnAssignmentTypes.Any())
                {
                    functionAsTypeNames.AddRange(functionReturnAssignmentTypes);
                }
            }

            return functionAsTypeNames;
        }

        private IEnumerable<string> AsTypeNameForValueParameters(string procedureIdentifier, ParserRuleContext procedureLetCtxt)
        {
            var propertyLets = _declarationFinderProvider.DeclarationFinder.MatchName(procedureIdentifier)
                .Where(f => f.DeclarationType.HasFlag(DeclarationType.PropertyLet)
                    && f.References.Any(rf => rf.Context.IsDescendentOf(procedureLetCtxt)))
                .Select(d => d as ModuleBodyElementDeclaration);

            return propertyLets.Select(d => d.Parameters.Last().AsTypeName);
        }

        private IEnumerable<string> AsTypeNameForFunctionTypes(string procedureIdentifier, ParserRuleContext procedureLetCtxt)
        {
            var functions = _declarationFinderProvider.DeclarationFinder.MatchName(procedureIdentifier)
                .Where(f => f.DeclarationType.HasFlag(DeclarationType.Function)
                    && f.References.Any(rf => rf.Context.IsDescendentOf(procedureLetCtxt)))
                .Select(d => d as ModuleBodyElementDeclaration);

            return functions.Select(f => f.AsTypeName);
        }

        public List<string> InferAsTypeNames(IReadOnlyCollection<VBAParser.LExprContext> tContexts)
        {
            var rhsDeclarationTypes = new List<DeclarationType>()
            {
                DeclarationType.Function,
                DeclarationType.Variable,
                DeclarationType.Constant
            };

            var members = _declarationFinderProvider.DeclarationFinder.Members(_target.QualifiedModuleName)
                .Where(d => d != _target || !d.DeclarationType.HasFlag(DeclarationType.Module));

            var nonMemberLocalRHSDeclarations = rhsDeclarationTypes.SelectMany(dt => _declarationFinderProvider.DeclarationFinder.UserDeclarations(dt))
                .Where(d => d.QualifiedModuleName != _target.QualifiedModuleName
                    && d.References.Any(rf => rf.QualifiedModuleName == _target.QualifiedModuleName));

            var externalAssignmentReferenceModules = _target.References
                .Where(rf => rf.IsAssignment && rf.QualifiedModuleName != _target.QualifiedModuleName)
                .Select(rf => rf.QualifiedModuleName).ToList();

            var externalRHSDeclarations = externalAssignmentReferenceModules.Any()
                ? GetExternalRHSAssignmentDeclarations(externalAssignmentReferenceModules, rhsDeclarationTypes)
                : Enumerable.Empty<Declaration>();

            var builtIns = _declarationFinderProvider.DeclarationFinder.AllBuiltInDeclarations
                .Where(f => f.DeclarationType.HasFlag(DeclarationType.Member)
                    && f.References.Any(rf => rf.QualifiedModuleName == _target.QualifiedModuleName
                        || externalAssignmentReferenceModules.Contains(rf.QualifiedModuleName)));

            var referenceIdentifierLExprPairs = members
                .Concat(nonMemberLocalRHSDeclarations)
                .Concat(externalRHSDeclarations)
                .Concat(builtIns)
                .SelectMany(d => d.References.Where(rf => !rf.IsAssignment))
                .Select(refID => (refID, refID.Context.GetAncestor<VBAParser.LExprContext>()))
                .ToList();

            return referenceIdentifierLExprPairs
                .Where(pr => tContexts
                .Contains(pr.Item2))
                .Select(lr => lr.refID.Declaration.AsTypeName)
                .ToList();
        }

        public List<string> InferAsTypeNames(IEnumerable<VBAParser.ArgumentListContext> argListContexts)
        {
            var parameterAsTypeNames = new List<string>();

            foreach (var argListContext in argListContexts)
            {
                (int? ZeroBasedParameterPosition, string ParameterName) = GetParameterLocatorAttributes(argListContext, _target.IdentifierName);

                if (!ZeroBasedParameterPosition.HasValue && string.IsNullOrEmpty(ParameterName))
                {
                    parameterAsTypeNames = new List<string>();
                    break;
                }
                var contextOfInterest = argListContext.GetAncestor<VBAParser.IndexExprContext>() as ParserRuleContext
                    ?? argListContext.GetAncestor<VBAParser.CallStmtContext>() as ParserRuleContext;

                var procedureIdentifier = contextOfInterest.children[0].GetText();

                var moduleBodyElement = _declarationFinderProvider.DeclarationFinder
                    .MatchName(procedureIdentifier)
                    .Where(d => d.DeclarationType.HasFlag(DeclarationType.Function)
                        || d.DeclarationType.HasFlag(DeclarationType.Procedure))
                    .SingleOrDefault(f => f.References.Any(rf => rf.Context.Parent == contextOfInterest))
                    as ModuleBodyElementDeclaration;

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

        private List<Declaration> GetExternalRHSAssignmentDeclarations(List<QualifiedModuleName> externalAssignmentReferenceModules, List<DeclarationType> declarationTypes)
        {
            return _target.Accessibility != Accessibility.Private 
                || _target.Accessibility != Accessibility.Implicit
                    ? externalAssignmentReferenceModules
                        .SelectMany(q => _declarationFinderProvider.DeclarationFinder.Members(q))
                        .ToList()
                    : new List<Declaration>();
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

        /// <summary>
        /// Wraps IParseTreeValueFactory coercing numeric AsTypeName results to a standard AsTypeName.
        /// Avoids generating declarations that will be flagged by inspections like IntegerDataTypeInspection.
        /// </summary>
        private struct LiteralExprContextToAsTypeNameConverter
        {
            private static Dictionary<string, string> _literalResultModifiers = new Dictionary<string, string>()
            {
                [Tokens.Byte] = Tokens.Long,
                [Tokens.Integer] = Tokens.Long,
                [Tokens.Single] = Tokens.Double,
            };

            private readonly IParseTreeValueFactory _parseTreeValueFactory;

            public LiteralExprContextToAsTypeNameConverter(IParseTreeValueFactory parseTreeValueFactory)
            {
                _parseTreeValueFactory = parseTreeValueFactory;
            }

            public string InferAsTypeName(VBAParser.LiteralExprContext literalExprContext)
            {
                var asTypeName = _parseTreeValueFactory?.Create(literalExprContext?.GetText()).ValueType ?? Tokens.Variant;

                return _literalResultModifiers.TryGetValue(asTypeName, out var modifiedAsTypeName)
                    ? modifiedAsTypeName
                    : asTypeName;
            }
        }
    }
}
