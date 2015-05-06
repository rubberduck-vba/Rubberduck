using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Listeners;

namespace Rubberduck.UI.Refactorings.ExtractMethod
{
    public class ExtractMethodRefactoring
    {
        public static IDictionary<VBAParser.AmbiguousIdentifierContext, ExtractedDeclarationUsage> GetParentMethodDeclarations(IParseTree parseTree, QualifiedSelection selection)
        {
            var declarations = parseTree.GetContexts<DeclarationListener, ParserRuleContext>(new DeclarationListener(selection.QualifiedName)).ToList();

            var constants = declarations.Select(d => d.Context).OfType<VBAParser.ConstSubStmtContext>().Select(constant => constant.ambiguousIdentifier());
            var variables = declarations.Select(d => d.Context).OfType<VBAParser.VariableSubStmtContext>().Select(variable => variable.ambiguousIdentifier());
            var arguments = declarations.Select(d => d.Context).OfType<VBAParser.ArgContext>().Select(arg => arg.ambiguousIdentifier());

            var identifiers = constants.Union(variables)
                                       .Union(arguments)
                                       .ToDictionary(declaration => declaration.GetText(),
                                                     declaration => declaration);

            var references = parseTree.GetContexts<VariableReferencesListener, VBAParser.AmbiguousIdentifierContext>(new VariableReferencesListener(selection.QualifiedName))
                                      .GroupBy(usage => new { Identifier = usage.Context.GetText() })
                                      .ToList();

            var notUsedInSelection = references.Where(usage => usage.All(token => !selection.Selection.Contains(token.Context.GetSelection())))
                                               .Select(usage => usage.Key).ToList();

            var usedBeforeSelection = references.Where(usage => usage.Any(token => token.Context.GetSelection().EndLine < selection.Selection.StartLine))
                                                    .Select(usage => usage.Key)
                                                .Where(usage => notUsedInSelection.All(e => e.Identifier != usage.Identifier));

            var usedAfterSelection = references.Where(usage => usage.Any(token => token.Context.GetSelection().StartLine > selection.Selection.EndLine))
                                                   .Select(usage => usage.Key)
                                                .Where(usage => notUsedInSelection.All(e => e.Identifier != usage.Identifier));

            var usedOnlyWithinSelection = references.Where(usage => usage.All(token => selection.Selection.Contains(token.Context.GetSelection())))
                                                    .Select(usage => usage.Key);


            var result = new Dictionary<VBAParser.AmbiguousIdentifierContext, ExtractedDeclarationUsage>();

            // temporal coupling: references used after selection must be added first
            foreach (var reference in usedAfterSelection)
            {
                VBAParser.AmbiguousIdentifierContext key;
                if (identifiers.TryGetValue(reference.Identifier, out key))
                {
                    if (!result.ContainsKey(key))
                    {
                        result.Add(key, ExtractedDeclarationUsage.UsedAfterSelection);
                    }
                }
            }

            foreach (var reference in usedBeforeSelection)
            {
                VBAParser.AmbiguousIdentifierContext key;
                if (identifiers.TryGetValue(reference.Identifier, out key))
                {
                    if (!result.ContainsKey(key))
                    {
                        result.Add(key, ExtractedDeclarationUsage.UsedBeforeSelection);
                    }
                }
            }

            foreach (var reference in usedOnlyWithinSelection)
            {
                VBAParser.AmbiguousIdentifierContext key;
                if (identifiers.TryGetValue(reference.Identifier, out key))
                {
                    if (!result.ContainsKey(key))
                    {
                        result.Add(key, ExtractedDeclarationUsage.UsedOnlyInSelection);
                    }
                }
            }

            foreach (var reference in notUsedInSelection)
            {
                VBAParser.AmbiguousIdentifierContext key;
                if (identifiers.TryGetValue(reference.Identifier, out key))
                {
                    if (!result.ContainsKey(key))
                    {
                        result.Add(key, ExtractedDeclarationUsage.NotUsed);
                    }
                }
            }

            return result;
        }
    }
}
