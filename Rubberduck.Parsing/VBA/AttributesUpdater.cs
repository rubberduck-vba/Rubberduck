using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using NLog;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.VBA
{
    public class AttributesUpdater : IAttributesUpdater
    {
        private readonly IParseTreeProvider _parseTreeProvider;
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();


        public AttributesUpdater(IParseTreeProvider parseTreeProvider)
        {
            _parseTreeProvider = parseTreeProvider;
        }


        public void AddAttribute(IRewriteSession rewriteSession, Declaration declaration, string attribute, IReadOnlyList<string> values)
        {
            if (string.IsNullOrEmpty(attribute))
            {
                return;
            }

            //Attributes must have at least one value.
            if (values == null || !values.Any())
            {
                return;
            }

            if (declaration == null)
            {
                _logger.Warn("Tried to add an attribute to a declaration that is null.");
                _logger.Trace($"Tried to add attribute {attribute} with values {AttributeValuesText(values)} to a declaration that is null.");
                return;
            }

            if (rewriteSession.TargetCodeKind != CodeKind.AttributesCode)
            {
                _logger.Warn($"Tried to add an attribute with a rewriter not suitable for attributes. (target code kind = {rewriteSession.TargetCodeKind})");
                _logger.Trace($"Tried to add attribute {attribute} with values {AttributeValuesText(values)} to {declaration.QualifiedModuleName} using a rewriter not suitable for attributes.");
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

            var rewriter = rewriteSession.CheckOutModuleRewriter(declaration.QualifiedModuleName);
            if (declaration.DeclarationType.HasFlag(DeclarationType.Module))
            {
                var codeToAdd = $"Attribute {attribute} = {AttributeValuesText(values)}";
                InsertAfterLastModuleAttribute(rewriter, declaration.QualifiedModuleName, codeToAdd);
            }
            else
            {
                var codeToAdd = $"Attribute {attribute} = {AttributeValuesText(values)}";
                InsertAfterFirstEolOfAttributeContext(rewriter, declaration, codeToAdd);
            }
        }

        private void InsertAfterLastModuleAttribute(IModuleRewriter rewriter, QualifiedModuleName module, string codeToAdd)
        {
            var moduleParseTree = (ParserRuleContext)_parseTreeProvider.GetParseTree(module, CodeKind.AttributesCode);
            var lastModuleAttribute = moduleParseTree.GetDescendents<VBAParser.ModuleAttributesContext>()
                .Where(moduleAttributes => moduleAttributes.attributeStmt() != null)
                .SelectMany(moduleAttributes => moduleAttributes.attributeStmt())
                .OrderBy(moduleAttribute => moduleAttribute.stop.TokenIndex)
                .LastOrDefault();
            if (lastModuleAttribute == null)
            {
                //This should never happen for a real module.
                var codeToInsert = codeToAdd + Environment.NewLine;
                rewriter.InsertBefore(0, codeToInsert);
            }
            else
            {
                var codeToInsert = Environment.NewLine + codeToAdd;
                rewriter.InsertAfter(lastModuleAttribute.stop.TokenIndex, codeToInsert);
            }
        }

        private void InsertAfterFirstEolOfAttributeContext(IModuleRewriter rewriter, Declaration declaration, string codeToAdd)
        {
            var attributesContext = declaration.AttributesPassContext;
            var firstEndOfLineInMember = attributesContext.GetDescendent<VBAParser.EndOfLineContext>();
            if (firstEndOfLineInMember == null)
            {
                var codeToInsert = Environment.NewLine + codeToAdd;
                rewriter.InsertAfter(declaration.AttributesPassContext.Stop.TokenIndex, codeToInsert);
            }
            else
            {
                var codeToInsert = codeToAdd + Environment.NewLine;
                rewriter.InsertAfter(firstEndOfLineInMember.Stop.TokenIndex, codeToInsert);
            }
        }

        private static string AttributeValuesText(IEnumerable<string> attributeValues)
        {
            return string.Join(", ", attributeValues);
        }

        public void RemoveAttribute(IRewriteSession rewriteSession, Declaration declaration, string attribute, IReadOnlyList<string> values = null)
        {
            if (string.IsNullOrEmpty(attribute))
            {
                return;
            }

            if (declaration == null)
            {
                _logger.Warn("Tried to remove an attribute from a declaration that is null.");
                _logger.Trace($"Tried to remove attribute {attribute} {(values != null ? $"with values {AttributeValuesText(values)} " : string.Empty)}from a declaration that is null.");
                return;
            }

            if (rewriteSession.TargetCodeKind != CodeKind.AttributesCode)
            {
                _logger.Warn($"Tried to remove an attribute with a rewriter not suitable for attributes. (target code kind = {rewriteSession.TargetCodeKind})");
                _logger.Trace($"Tried to remove attribute {attribute} {(values != null ? $"with values {AttributeValuesText(values)} " : string.Empty)}from {declaration.QualifiedModuleName} using a rewriter not suitable for attributes.");
                return;
            }

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
            if (string.IsNullOrEmpty(attribute))
            {
                return;
            }

            //Attributes must have at least one value.
            if (newValues == null || !newValues.Any())
            {
                return;
            }

            if (declaration == null)
            {
                _logger.Warn("Tried to updtae an attribute on a declaration that is null.");
                _logger.Trace($"Tried to update values for attribute {attribute} {(oldValues != null ? $"with values {AttributeValuesText(oldValues)} " : string.Empty)}on a declaration that is null.");
                return;
            }

            if (rewriteSession.TargetCodeKind != CodeKind.AttributesCode)
            {
                _logger.Warn($"Tried to update an attribute with a rewriter not suitable for attributes. (target code kind = {rewriteSession.TargetCodeKind})");
                _logger.Trace($"Tried to update values for attribute {attribute} {(oldValues != null ? $"with values {AttributeValuesText(oldValues)} " : string.Empty)}on {declaration.QualifiedModuleName} to {AttributeValuesText(oldValues)} using a rewriter not suitable for attributes.");
                return;
            }

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
            var replacementText = $" {AttributeValuesText(values)}";

            rewriter.Replace(new Interval(firstIndexToReplace, lastIndexToReplace), replacementText);
        }

        public void AddOrUpdateAttribute(
            IRewriteSession rewriteSession, 
            Declaration declaration, 
            string attribute,
            IReadOnlyList<string> values)
        {
            var attributeNodes = ApplicableAttributeNodes(declaration, attribute);

            if (!attributeNodes.Any())
            {
                AddAttribute(rewriteSession, declaration, attribute, values);
                return;
            }

            if (attribute.Equals("VB_Ext_Key"))
            {
                var newKey = values.First();
                var matchingExtKeyAttribute = attributeNodes.FirstOrDefault(node =>  newKey.Equals(node.Values.FirstOrDefault(), StringComparison.InvariantCultureIgnoreCase));

                if (matchingExtKeyAttribute == null)
                {
                    AddAttribute(rewriteSession, declaration, attribute, values);
                    return;
                }

                var oldValues = matchingExtKeyAttribute.Values;
                UpdateAttribute(rewriteSession, declaration, attribute, values, oldValues);
                return;
            }

            UpdateAttribute(rewriteSession, declaration, attribute, values);
        }
    }
}