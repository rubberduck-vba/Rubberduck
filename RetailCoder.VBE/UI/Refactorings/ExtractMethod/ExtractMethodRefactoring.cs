using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Rubberduck.Extensions;
using Rubberduck.VBA;
using Rubberduck.VBA.Grammar;
using Rubberduck.VBA.ParseTreeListeners;

namespace Rubberduck.UI.Refactorings.ExtractMethod
{
    public class ExtractMethodRefactoring
    {
        public static IDictionary<VBParser.AmbiguousIdentifierContext, ExtractedDeclarationUsage> GetParentMethodDeclarations(IParseTree parseTree, QualifiedSelection selection)
        {
            var declarations = parseTree.GetContexts<DeclarationListener, ParserRuleContext>(new DeclarationListener(selection.QualifiedName)).ToList();

            var constants = declarations.Select(d => d.Context).OfType<VBParser.ConstSubStmtContext>().Select(constant => constant.AmbiguousIdentifier());
            var variables = declarations.Select(d => d.Context).OfType<VBParser.VariableSubStmtContext>().Select(variable => variable.AmbiguousIdentifier());
            var arguments = declarations.Select(d => d.Context).OfType<VBParser.ArgContext>().Select(arg => arg.AmbiguousIdentifier());

            var identifiers = constants.Union(variables)
                                       .Union(arguments)
                                       .ToDictionary(declaration => declaration.GetText(),
                                                     declaration => declaration);

            var references = parseTree.GetContexts<VariableReferencesListener, VBParser.AmbiguousIdentifierContext>(new VariableReferencesListener(selection.QualifiedName))
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


            var result = new Dictionary<VBParser.AmbiguousIdentifierContext, ExtractedDeclarationUsage>();

            // temporal coupling: references used after selection must be added first
            foreach (var reference in usedAfterSelection)
            {
                VBParser.AmbiguousIdentifierContext key;
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
                VBParser.AmbiguousIdentifierContext key;
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
                VBParser.AmbiguousIdentifierContext key;
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
                VBParser.AmbiguousIdentifierContext key;
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
