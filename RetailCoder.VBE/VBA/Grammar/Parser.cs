using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Vbe.Interop;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.VBA.Grammar
{
    [ComVisible(false)]
    public class Parser
    {
        private readonly IEnumerable<ISyntax> _grammar;
        public Parser(IEnumerable<ISyntax> grammar)
        {
            _grammar = grammar;
        }

        public SyntaxTreeNode Parse(VBProject project)
        {
            var nodes = new List<SyntaxTreeNode>();
            try
            {
                var components = project.VBComponents.Cast<VBComponent>().ToList();
                foreach (var component in components)
                {
                    var lineCount = component.CodeModule.CountOfLines;
                    if (lineCount <= 0)
                    {
                        continue;
                    }

                    var code = component.CodeModule.Lines[1, lineCount];
                    var isClassModule = component.Type == vbext_ComponentType.vbext_ct_ClassModule
                                        || component.Type == vbext_ComponentType.vbext_ct_Document
                                        || component.Type == vbext_ComponentType.vbext_ct_MSForm;

                    nodes.Add(Parse(project.Name, component.Name, code, isClassModule));
                }
            }
            catch
            {
                // todo: handle exception like a chief
                Debug.Assert(false);
            }

            return new ProjectNode(project, nodes);
        }

        /// <summary>
        /// Converts VBA code into a <see cref="SyntaxTreeNode"/>.
        /// </summary>
        /// <param name="projectName">The name of the VBA Project, used for scoping public nodes.</param>
        /// <param name="componentName">The name of the module, used for scoping private nodes.</param>
        /// <param name="code">The code to parse.</param>
        /// <returns></returns>
        public SyntaxTreeNode Parse(string projectName, string componentName, string code, bool isClassModule)
        {
            var content = SplitLogicalCodeLines(projectName, componentName, code);
            var memberNodes = ParseModuleMembers(projectName, componentName, content).ToList();

            var result = new ModuleNode(projectName, componentName, memberNodes, isClassModule);
            return result;
        }

        private IEnumerable<LogicalCodeLine> SplitLogicalCodeLines(string projectName, string componentName, string content)
        {
            const string lineContinuationMarker = "_";

            var lines = content.Split('\n').Select(line => line.Replace("\r", string.Empty)).ToList();

            var logicalLine = new StringBuilder();
            var startLine = 0;
            var isContinuing = false;
            for (var index = 0; index < lines.Count; index++)
            {
                if (!isContinuing)
                {
                    startLine = index + 1;
                }

                var line = lines[index];
                if (line.EndsWith(lineContinuationMarker))
                {
                    isContinuing = true;
                    logicalLine.Append(line.Remove(line.Length - 1));
                }
                else
                {
                    logicalLine.Append(line);
                    yield return new LogicalCodeLine(projectName, componentName, startLine, index + 1, logicalLine.ToString());
                    logicalLine.Clear();
                    isContinuing = false;
                }
            }
        }

        private IEnumerable<SyntaxTreeNode> ParseModuleMembers(string publicScope, string localScope, IEnumerable<LogicalCodeLine> logicalCodeLines)
        {
            var lines = logicalCodeLines.ToArray();
            for (var index = 0; index < lines.Length; index++)
            {
                var line = lines[index];
                if (string.IsNullOrWhiteSpace(line.Content))
                {
                    continue;
                }

                var instructions = line.SplitInstructions();
                foreach (var instruction in instructions)
                {
                    var parsed = false;
                    foreach (var syntax in _grammar.Where(s => !s.IsChildNodeSyntax))
                    {
                        SyntaxTreeNode node;
                        if (!syntax.IsMatch(publicScope, localScope, instruction, out node))
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
                                    yield return ParseProcedure(publicScope, localScope + "." + (node as ProcedureNode).Identifier.Name, node as ProcedureNode, lines, ref index);
                                    parsed = true;
                                    break;
                                }

                                yield return ParseCodeBlock(publicScope, localScope, codeBlockNode, lines, ref index);
                                parsed = true;
                                break;
                            }
                        }
                        
                        yield return node;
                        parsed = true;
                    }

                    if (!parsed)
                    {
                        yield return new ExpressionNode(instruction, localScope);
                    }
                }
            }
        }

        private SyntaxTreeNode ParseCodeBlock(string publicScope, string localScope, CodeBlockNode codeBlockNode, IEnumerable<LogicalCodeLine> logicalLines, ref int index)
        {
            var ifBlockNode = codeBlockNode as IfBlockNode;
            if (ifBlockNode != null && !string.IsNullOrEmpty(ifBlockNode.Expression.Value))
            {
                return codeBlockNode;
            }

            var result = codeBlockNode;
            var grammar = result.ChildSyntaxType == null
                ? _grammar.Where(syntax => !syntax.IsChildNodeSyntax).ToList()
                : _grammar.Where(syntax => syntax.IsChildNodeSyntax && syntax.GetType() == result.ChildSyntaxType).ToList();

            var logicalCodeLines = logicalLines as LogicalCodeLine[] ?? logicalLines.ToArray();
            var lines = logicalCodeLines.ToArray();

            var currentIndex = ++index;
            while (currentIndex < lines.Length && !result.EndOfBlockMarkers.Any(marker => lines[currentIndex].Content.Trim().StartsWith(marker)))
            {
                var line = lines[currentIndex];
                if (string.IsNullOrWhiteSpace(line.Content))
                {
                    currentIndex++;
                    continue;
                }

                var instructions = line.SplitInstructions();
                foreach (var instruction in instructions)
                {
                    var parsed = false;
                    foreach (var syntax in grammar)
                    {
                        SyntaxTreeNode node;
                        if (!syntax.IsMatch(publicScope, localScope, instruction, out node))
                        {
                            continue;
                        }

                        var childNode = node as CodeBlockNode;
                        if (childNode != null)
                        {
                            node = ParseCodeBlock(publicScope, localScope, childNode, logicalCodeLines, ref currentIndex);
                        }

                        result.AddNode(node);
                        parsed = true;
                        break;
                    }

                    if (!parsed)
                    {
                        result.AddNode(new ExpressionNode(instruction, localScope));
                    }
                }

                if (lines[currentIndex + 1].Content.Trim().StartsWith(ReservedKeywords.Else))
                {
                    break;
                }

                currentIndex++;
            }

            index = currentIndex;
            return result;
        }

        private SyntaxTreeNode ParseProcedure(string publicScope, string localScope, ProcedureNode procedureNode, IEnumerable<LogicalCodeLine> logicalLines, ref int index)
        {
            var result = procedureNode;
            var grammar = VBAGrammar.GetGrammarSyntax().Where(s => !s.IsChildNodeSyntax).ToList();

            var logicalCodeLines = logicalLines as LogicalCodeLine[] ?? logicalLines.ToArray();
            var lines = logicalCodeLines.ToArray();

            var currentIndex = ++index;
            while (currentIndex < lines.Length && !result.EndOfBlockMarkers.Any(marker => lines[currentIndex].Content.Trim().StartsWith(marker)))
            {
                var line = lines[currentIndex];
                if (string.IsNullOrWhiteSpace(line.Content))
                {
                    currentIndex++;
                    continue;
                }

                var instructions = line.SplitInstructions();
                foreach (var instruction in instructions)
                {
                    var parsed = false;
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
                                node = ParseCodeBlock(publicScope, localScope, childNode, logicalCodeLines, ref currentIndex);
                            }
                        }

                        result.AddNode(node);
                        parsed = true;
                        break;
                    }

                    if (!parsed)
                    {
                        result.AddNode(new ExpressionNode(instruction, localScope));
                    }
                }
                currentIndex++;
            }

            index = currentIndex;
            return result;
        }
    }
}
