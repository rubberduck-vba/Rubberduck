using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Rubberduck.Reflection;
using Rubberduck.VBA.Parser.Grammar;

namespace Rubberduck.VBA.Parser
{
    [ComVisible(false)]
    public class Parser
    {
        /// <summary>
        /// Converts VBA code into a <see cref="SyntaxTreeNode"/>.
        /// </summary>
        /// <param name="projectName">The name of the VBA Project, used for scoping public nodes.</param>
        /// <param name="componentName">The name of the module, used for scoping private nodes.</param>
        /// <param name="code">The code to parse.</param>
        /// <returns></returns>
        public SyntaxTreeNode Parse(string projectName, string componentName, string code)
        {
            try
            {
                var content = SplitLogicalCodeLines(projectName, componentName, code);
                var memberNodes = ParseModuleMembers(projectName, componentName, content).ToList();

                var result = new ModuleNode(projectName, componentName, memberNodes);

                return result;

            }
            catch (Exception exception)
            {             
                throw;
            }
        }

        private IEnumerable<LogicalCodeLine> SplitLogicalCodeLines(string projectName, string componentName, string content)
        {
            const string lineContinuationMarker = "_";

            var lines = content.Split('\n').Select(line => line.Replace("\r", string.Empty)).ToList();

            var logicalLine = new StringBuilder();
            for (int index = 0; index < lines.Count; index++)
            {
                var line = lines[index];
                if (line.EndsWith(lineContinuationMarker))
                {
                    logicalLine.Append(line.Remove(line.Length - 1));
                }
                else
                {
                    logicalLine.Append(line);
                    yield return new LogicalCodeLine(projectName, componentName, index + 1, logicalLine.ToString());
                    logicalLine.Clear();
                }
            }
        }

        private IEnumerable<SyntaxTreeNode> ParseModuleMembers(string publicScope, string localScope, IEnumerable<LogicalCodeLine> logicalCodeLines)
        {
            var currentLocalScope = localScope;
            var lines = logicalCodeLines.ToArray();
            for (var index = 0; index < lines.Length; index++)
            {
                var line = lines[index];
                var instructions = line.SplitInstructions();
                foreach (var instruction in instructions)
                {
                    foreach (var syntax in VBAGrammar.GetGrammarSyntax().Where(s => !s.IsChildNodeSyntax))
                    {
                        SyntaxTreeNode node;
                        if (!syntax.IsMatch(publicScope, currentLocalScope, instruction, out node))
                        {
                            continue;
                        }

                        if (syntax.Type.HasFlag(SyntaxType.HasChildNodes))
                        {
                            var codeBlockNode = node as CodeBlockNode;
                            if (codeBlockNode != null)
                            {
                                if (node is ProcedureNode)
                                {
                                    currentLocalScope = localScope + "." + (node as ProcedureNode).Identifier.Name;
                                    yield return  ParseProcedure(publicScope, currentLocalScope, node as ProcedureNode, lines, ref index);
                                    continue;
                                }

                                yield return ParseCodeBlock(publicScope, currentLocalScope, codeBlockNode, lines, ref index);
                                currentLocalScope = localScope;
                                continue;
                            }
                        }

                        yield return node;
                    }
                }
            }
        }

        private SyntaxTreeNode ParseCodeBlock(string publicScope, string localScope, CodeBlockNode codeBlockNode, IEnumerable<LogicalCodeLine> logicalLines, ref int index)
        {
            var result = codeBlockNode;
            var grammar = result.ChildSyntaxType == null
                ? VBAGrammar.GetGrammarSyntax().ToList()
                : VBAGrammar.GetGrammarSyntax().Where(syntax => syntax.IsChildNodeSyntax && syntax.GetType() == result.ChildSyntaxType).ToList();

            var logicalCodeLines = logicalLines as LogicalCodeLine[] ?? logicalLines.ToArray();
            var lines = logicalCodeLines.ToArray();
            index++;

            while (index < lines.Length && !result.EndOfBlockMarkers.Contains(lines[index].Content.Trim()))
            {
                var line = lines[index];
                var instructions = line.SplitInstructions();
                foreach (var instruction in instructions)
                {
                    foreach (var syntax in grammar)
                    {
                        SyntaxTreeNode node;
                        if (!syntax.IsMatch(publicScope, localScope, instruction, out node))
                        {
                            continue;
                        }

                        if (node.HasChildNodes)
                        {
                            var childNode = node as CodeBlockNode;
                            if (childNode != null)
                            {
                                node = ParseCodeBlock(publicScope, localScope, childNode, logicalCodeLines, ref index);
                            }
                        }

                        result = result.AddNode<CodeBlockNode>(node);
                        break;
                    }
                }
                index++;
            }

            return result;
        }

        private SyntaxTreeNode ParseProcedure(string publicScope, string localScope, ProcedureNode procedureNode, IEnumerable<LogicalCodeLine> logicalLines, ref int index)
        {
            var result = procedureNode;
            var grammar = result.ChildSyntaxType == null
                ? VBAGrammar.GetGrammarSyntax().ToList()
                : VBAGrammar.GetGrammarSyntax().Where(syntax => syntax.IsChildNodeSyntax && syntax.GetType() == result.ChildSyntaxType).ToList();

            var logicalCodeLines = logicalLines as LogicalCodeLine[] ?? logicalLines.ToArray();
            var lines = logicalCodeLines.ToArray();
            index++;

            while (index < lines.Length && !result.EndOfBlockMarkers.Contains(lines[index].Content.Trim()))
            {
                var line = lines[index];
                var instructions = line.SplitInstructions();
                foreach (var instruction in instructions)
                {
                    foreach (var syntax in grammar)
                    {
                        SyntaxTreeNode node;
                        if (!syntax.IsMatch(publicScope, localScope, instruction, out node)) continue;

                        if (node.HasChildNodes)
                        {
                            var childNode = node as CodeBlockNode;
                            if (childNode != null)
                            {
                                node = ParseCodeBlock(publicScope, localScope, childNode, logicalCodeLines, ref index);
                            }
                        }

                        result = result.AddNode<ProcedureNode>(node);
                        break;
                    }
                }
                index++;
            }

            return result;
        }
    }
}
