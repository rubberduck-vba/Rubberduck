using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA.Parsing;

namespace Rubberduck.Parsing.VBA
{
    public class AttributesUpdater : IAttributesUpdater
    {
        private readonly IParseTreeProvider _parseTreeProvider;

        public AttributesUpdater(IParseTreeProvider parseTreeProvider)
        {
            _parseTreeProvider = parseTreeProvider;
        }


        public void AddAttribute(IRewriteSession rewriteSession, Declaration declaration, string attribute, IReadOnlyList<string> values)
        {
            Debug.Assert(rewriteSession.TargetCodeKind == CodeKind.AttributesCode);

            //Attributes must have at least one value.
            if (values == null || !values.Any())
            {
                return;
            }

            //VB_Ext_Key is special in that this attribute can be declared multiple times, but only once for each key.
            if (attribute.ToUpperInvariant().EndsWith("VB_EXT_KEY"))
            {
                if (values.Count != 2)
                {
                    return;
                }
                //Is the key already defined as external key?
                if(declaration.Attributes.Any(attrbt => attrbt.Name.Equals(attribute, StringComparison.OrdinalIgnoreCase)
                                                        && attrbt.Values[0].Equals(values[0])))
                {
                    return;
                }
            }
            else if (declaration.Attributes.HasAttribute(attribute))
            {
                return;
            }

            var codeToInsert = $"{Environment.NewLine}Attribute {attribute} ={AttributeValuesText(values)}";

            var rewriter = rewriteSession.CheckOutModuleRewriter(declaration.QualifiedModuleName);
            if (declaration.DeclarationType.HasFlag(DeclarationType.Module))
            {
                var moduleParseTree = (ParserRuleContext)_parseTreeProvider.GetParseTree(declaration.QualifiedModuleName, CodeKind.AttributesCode);
                var lastModuleAttribute = moduleParseTree.GetDescendents<VBAParser.ModuleAttributesContext>()
                    .Where(moduleAttributes => moduleAttributes.attributeStmt() != null)
                    .SelectMany(moduleAttributes => moduleAttributes.attributeStmt())
                    .OrderBy(moduleAttribute => moduleAttribute.stop.TokenIndex)
                    .Last();
                rewriter.InsertAfter(lastModuleAttribute.stop.TokenIndex, codeToInsert);
            }
            else
            {
                rewriter.InsertAfter(declaration.AttributesPassContext.Stop.TokenIndex, codeToInsert);
            }
        }

        private static string AttributeValuesText(IEnumerable<string> attributeValues)
        {
            var builder = new StringBuilder();
            foreach (var attributeValue in attributeValues)
            {
                builder.Append($" {attributeValue},");
            }
            //Remove trailing comma.
            builder.Length--;
            return builder.ToString();
        }

        public void RemoveAttribute(IRewriteSession rewriteSession, Declaration declaration, string attribute, IReadOnlyList<string> values = null)
        {
            Debug.Assert(rewriteSession.TargetCodeKind == CodeKind.AttributesCode);

            var attributeNodes = ApplicableAttributeNodes(declaration, attribute, values);

            if (!attributeNodes.Any())
            {
                return;
            }

            var rewriter = rewriteSession.CheckOutModuleRewriter(declaration.QualifiedModuleName);
            RemoveNodes(rewriter, attributeNodes);
        }

        private static void RemoveNodes(IModuleRewriter rewriter, IEnumerable<AttributeNode> attributeNodes)
        {
            foreach (var node in attributeNodes)
            {
                var attributeContext = node.Context;
                rewriter.Remove(attributeContext);
                if (attributeContext.TryGetFollowingContext(out VBAParser.EndOfLineContext followingEndOfLine))
                {
                    rewriter.Remove(followingEndOfLine);
                }
                else if (attributeContext.TryGetPrecedingContext(out VBAParser.EndOfLineContext precedingEndOfLine))
                {
                    //We are on the last line. So, we must remove the preceding newline.
                    rewriter.Remove(precedingEndOfLine);
                }
            }
        }

        private static IList<AttributeNode> ApplicableAttributeNodes(Declaration declaration, string attribute, IReadOnlyList<string> values = null)
        {
            var attributeNodes = declaration.Attributes.AttributeNodes(attribute);
            if (values != null)
            {
                attributeNodes = attributeNodes.Where(node => node.Values.SequenceEqual(values));
            }

            return attributeNodes.ToList();
        }

        public void UpdateAttribute(IRewriteSession rewriteSession, Declaration declaration, string attribute, IReadOnlyList<string> newValues, IReadOnlyList<string> oldValues = null)
        {
            Debug.Assert(rewriteSession.TargetCodeKind == CodeKind.AttributesCode);

            var attributeNodes = ApplicableAttributeNodes(declaration, attribute, oldValues);

            if (!attributeNodes.Any())
            {
                return;
            }

            var rewriter = rewriteSession.CheckOutModuleRewriter(declaration.QualifiedModuleName);

            var nodeToUpdate = attributeNodes[0];
            UpdateAttributeValues(rewriter, nodeToUpdate, newValues);
            RemoveNodes(rewriter, attributeNodes.Skip(1));
        }

        private static void UpdateAttributeValues(IModuleRewriter rewriter, AttributeNode nodeToUpdate, IEnumerable<string> values)
        {
            var statementContext = nodeToUpdate.Context;
            var firstIndexToReplace = statementContext.EQ().Symbol.TokenIndex + 1;
            var lastIndexToReplace = statementContext.stop.TokenIndex;
            var replacementText = AttributeValuesText(values);

            rewriter.Replace(new Interval(firstIndexToReplace, lastIndexToReplace), replacementText);
        }
    }
}