using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.ImplicitTypeToExplicit
{
    internal class AsTypeNamesResultsHandler
    {
        private static List<string> _numericTypePrecedence = new List<string>()
        {
            Tokens.Double,
            Tokens.Single,
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
            Tokens.Decimal,
            Tokens.LongLong
        };

        private Dictionary<string, HashSet<string>> _candidates;

        public AsTypeNamesResultsHandler()
        {
            _candidates = new Dictionary<string, HashSet<string>>();
        }

        public void AddCandidates(Dictionary<string, List<string>> newCandidates)
        {
            foreach (var key in newCandidates.Keys)
            {
                AddCandidates(key, newCandidates[key]);
            }
        }

        public void AddCandidates(string identifier, IReadOnlyCollection<string> asTypeNames)
        {
            if (!asTypeNames.Any() || asTypeNames.All(s => string.IsNullOrEmpty(s)))
            {
                return;
            }

            if (!_candidates.ContainsKey(identifier))
            {
                _candidates.Add(identifier, new HashSet<string>());
            }

            foreach (var asTypeName in asTypeNames)
            {
                _candidates[identifier].Add(asTypeName);
            }
        }

        public void AddIndeterminantResult() 
            => AddCandidates(nameof(ParserRuleContext), new List<string>() { Tokens.Variant });

        public string ResolveAsTypeName(Declaration target)
        {
            if (!_candidates.Any())
            {
                return DefaultToVariantOrObject(target);
            }

            var nonLiteralExpressionAsTypeNames = new List<string>();
            foreach (var key in _candidates.Keys.Where(k => k != nameof(VBAParser.LiteralExprContext)))
            {
                nonLiteralExpressionAsTypeNames.AddRange(_candidates[key]);
            }

            var resolvedAsTypeName = ResolveAsTypeName(nonLiteralExpressionAsTypeNames);

            resolvedAsTypeName = ResolveAsTypeNameWithLiteralExpressionInferences(resolvedAsTypeName, _candidates);

            return resolvedAsTypeName
                ?? (_candidates.TryGetValue(nameof(VBAParser.LetStmtContext), out var asTypes) && asTypes.Any()
                    ? Tokens.Variant
                    : null)
                ?? DefaultToVariantOrObject(target);
        }

        private static string ResolveAsTypeName(List<string> asTypeNames)
        {
            return asTypeNames.Count() == 1
                ? asTypeNames.ElementAt(0) //Only one kind of AsTypeName found..the happy path
                : asTypeNames.Any(n => _allOrNothingTypeNames.Contains(n))
                    || asTypeNames.Any(n => !_numericTypePrecedence.Contains(n))
                        ? null //Indeterminant from resolved AsTypeNames
                        : _numericTypePrecedence.Find(t => asTypeNames.Contains(t)); //Use best numeric AsTypeName
        }

        private static string ResolveAsTypeNameWithLiteralExpressionInferences(string resolvedAsTypeName, Dictionary<string, HashSet<string>> asTypeNames)
        {
            if (!asTypeNames.ContainsKey(nameof(VBAParser.LiteralExprContext)))
            {
                return resolvedAsTypeName;
            }

            var resolvedLiteralsTypeName = ResolveAsTypeName(asTypeNames[nameof(VBAParser.LiteralExprContext)].ToList());
            if ((resolvedLiteralsTypeName != null)
                && (resolvedAsTypeName != null)
                && _numericTypePrecedence.Contains(resolvedLiteralsTypeName)
                && _numericTypePrecedence.Contains(resolvedAsTypeName))
            {
                if (_numericTypePrecedence.IndexOf(resolvedLiteralsTypeName) < _numericTypePrecedence.IndexOf(resolvedAsTypeName)
                    && resolvedLiteralsTypeName.Equals(Tokens.Double))
                {
                    return Tokens.Double;
                }
            }
            return resolvedAsTypeName ?? resolvedLiteralsTypeName;
        }

        private static string DefaultToVariantOrObject(Declaration target)
        {
            if (!target.References.Any())
            {
                return Tokens.Variant;
            }

            var isUsedInLetStmtLHS = target.References.Any(rf => rf.IsAssignment && !rf.IsSetAssignment);

            var isUsedAsLetStmtRHS = target.References.Select(rf => rf.Context).Any(c => c.Parent.Parent is VBAParser.LetStmtContext letStmt
                && letStmt.children.Last() == c.Parent);

            return isUsedInLetStmtLHS || isUsedAsLetStmtRHS
                ? Tokens.Variant
                : Tokens.Object;
        }
    }
}
