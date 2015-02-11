using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Antlr4.Runtime.Tree;
using Rubberduck.Extensions;
using Rubberduck.VBA;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.UI.Refactorings.ExtractMethod
{
    [ComVisible(false)]
    public class ExtractMethodRefactoring
    {
        public static IDictionary<VisualBasic6Parser.AmbiguousIdentifierContext, ExtractedDeclarationUsage> GetParentMethodDeclarations(IParseTree parseTree, Selection selection)
        {
            var declarations = parseTree.GetDeclarations().ToList();

            var constants = declarations.OfType<VisualBasic6Parser.ConstSubStmtContext>().Select(constant => constant.ambiguousIdentifier());
            var variables = declarations.OfType<VisualBasic6Parser.VariableSubStmtContext>().Select(variable => variable.ambiguousIdentifier());
            var arguments = declarations.OfType<VisualBasic6Parser.ArgContext>().Select(arg => arg.ambiguousIdentifier());

            var identifiers = constants.Union(variables)
                                       .Union(arguments)
                                       .ToDictionary(declaration => declaration.GetText(), 
                                                     declaration => declaration);

            var references = parseTree.GetVariableReferences()
                                      .GroupBy(usage => new { Identifier = usage.GetText()})
                                      .ToList();

            var notUsedInSelection = references.Where(usage => usage.All(token => !selection.Contains(token.GetSelection())))
                                               .Select(usage => usage.Key).ToList();

            var usedBeforeSelection = references.Where(usage => usage.Any(token => token.GetSelection().EndLine < selection.StartLine))
                                                    .Select(usage => usage.Key)
                                                .Where(usage => notUsedInSelection.All(e => e.Identifier != usage.Identifier));

            var usedAfterSelection = references.Where(usage => usage.Any(token => token.GetSelection().StartLine > selection.EndLine))
                                                   .Select(usage => usage.Key)
                                                .Where(usage => notUsedInSelection.All(e => e.Identifier != usage.Identifier));

            var usedOnlyWithinSelection = references.Where(usage => usage.All(token => selection.Contains(token.GetSelection())))
                                                    .Select(usage => usage.Key);


            var result = new Dictionary<VisualBasic6Parser.AmbiguousIdentifierContext, ExtractedDeclarationUsage>();

            // temporal coupling: references used after selection must be added first
            foreach (var reference in usedAfterSelection)
            {
                VisualBasic6Parser.AmbiguousIdentifierContext key;
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
                VisualBasic6Parser.AmbiguousIdentifierContext key;
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
                VisualBasic6Parser.AmbiguousIdentifierContext key;
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
                VisualBasic6Parser.AmbiguousIdentifierContext key;
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
