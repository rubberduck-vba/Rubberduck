using Rubberduck.Reflection.VBA.Grammar;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using Microsoft.Vbe.Interop;
using System.Text.RegularExpressions;

namespace Rubberduck.Reflection.VBA
{
    internal class Parser
    {
        private IEnumerable<ISyntax> _grammar;

        public Parser()
        {
            var syntaxType = typeof(ISyntax);
            _grammar = Assembly.GetExecutingAssembly()
                               .GetTypes()
                               .Where(type => type.BaseType == typeof(SyntaxBase))
                               .Select(type => type.GetConstructor(Type.EmptyTypes).Invoke(Type.EmptyTypes))
                               .Cast<ISyntax>()
                               .ToList();
        }

        public SyntaxTreeNode Parse(CodeModule module)
        {
            var project = module.Parent.Collection.Parent;
            var component = module.Parent;

            var publicScope = project.Name;
            var localScope = string.Concat(project.Name, ".", component.Name);

            var content = SplitLogicalCodeLines(project.Name, component.Name, module.Lines[1, module.CountOfLines]);
            var memberNodes = ParseModuleMembers(publicScope, localScope, content).ToList();

            var result = new ModuleNode(project.Name, component.Name, memberNodes);

            return result;
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
            for (int index = 0; index < lines.Length; index++)
            {
                var line = lines[index];
                var instructions = line.SplitInstructions();
                foreach (var instruction in instructions)
                {
                    SyntaxTreeNode node;
                    foreach (var syntax in _grammar.Where(syntax => !syntax.IsChildNodeSyntax))
                    {
                        if (syntax.IsMatch(publicScope, localScope, instruction, out node))
                        {
                            if (node.HasChildNodes)
                            {
                                var codeBlockNode = node as CodeBlockNode;
                                if (codeBlockNode != null)
                                {
                                    node = ParseCodeBlock(publicScope, localScope, codeBlockNode, lines, ref index);
                                }
                                else
                                {
                                    var declarationNode = node as DeclarationNodeBase;
                                    if (declarationNode != null)
                                    {
                                        //node = ParseDeclaration(publicScope, localScope, instruction);
                                    }
                                }
                            }
                            yield return node;
                            break;
                        }
                    }
                }
            }
        }
        
        private CodeBlockNode ParseCodeBlock(string publicScope, string localScope, CodeBlockNode codeBlockNode, IEnumerable<LogicalCodeLine> logicalLines, ref int index)
        {
            var result = codeBlockNode;
            var grammar = codeBlockNode.ChildSyntaxType == null 
                ? _grammar 
                : _grammar.Where(syntax => syntax.IsChildNodeSyntax && syntax.GetType() == codeBlockNode.ChildSyntaxType);

            var lines = logicalLines.ToArray();

            while (codeBlockNode.EndOfBlockMarkers.Contains(lines[index].Content.Trim()))
            {
                var line = lines[index];
                var instructions = line.SplitInstructions();
                foreach (var instruction in instructions)
                {
                    SyntaxTreeNode node;
                    foreach (var syntax in grammar)
                    {
                        if (syntax.IsMatch(publicScope, localScope, instruction, out node))
                        {
                            if (node.HasChildNodes)
                            {
                                var childNode = node as CodeBlockNode;
                                node = ParseCodeBlock(publicScope, localScope, childNode, logicalLines, ref index);
                            }
                            result = codeBlockNode.AddNode(node);
                            break;
                        }
                    }
                }
                index++;
            }

            return result;
        }
    }
}
