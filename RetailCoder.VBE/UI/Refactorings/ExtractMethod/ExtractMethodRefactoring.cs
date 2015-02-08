using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Microsoft.Vbe.Interop;
using Rubberduck.Extensions;
using Rubberduck.Refactoring;
using Rubberduck.VBA;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.UI.Refactorings.ExtractMethod
{
    [ComVisible(false)]
    public class ExtractMethodRefactoring : IRefactoring
    {
        public void Refactor(CodeModule module)
        {
            throw new NotImplementedException();
        }

        public static IDictionary<VisualBasic6Parser.AmbiguousIdentifierContext, ExtractedDeclarationUsage> GetParentMethodDeclarations(IParseTree parseTree, Selection selection)
        {
            var declarations = parseTree.GetDeclarations().ToList();
            var constants = declarations.OfType<VisualBasic6Parser.ConstSubStmtContext>().Select(constant => constant.ambiguousIdentifier());
            var variables = declarations.OfType<VisualBasic6Parser.VariableSubStmtContext>().Select(variable => variable.ambiguousIdentifier());
            
            var identifiers = constants.Union(variables)
                                       .ToDictionary(declaration => declaration.GetText(), 
                                                     declaration => declaration);

            var references = parseTree.GetVariableReferences()
                                      .GroupBy(usage => new { Identifier = usage.GetText()})
                                      .ToList();

            var usedBeforeSelection = references.Where(usage => usage.Any(token => token.GetSelection().EndLine < selection.StartLine))
                                                    .Select(usage => usage.Key);

            var usedAfterSelection = references.Where(usage => usage.Any(token => token.GetSelection().StartLine > selection.EndLine))
                                                   .Select(usage => usage.Key);

            var usedOnlyWithinSelection = references.Where(usage => usage.All(token => selection.Contains(token.GetSelection())))
                                                    .Select(usage => usage.Key);

            var notUsedInSelection = references.Where(usage => usage.All(token => !selection.Contains(token.GetSelection())))
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
