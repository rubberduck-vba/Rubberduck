using RetailCoderVBE.Reflection.VBA.Grammar;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using Microsoft.Vbe.Interop;

namespace RetailCoderVBE.Reflection.VBA
{
    internal class Parser
    {
        private IEnumerable<ISyntax> _grammar;

        public Parser()
        {
            var syntaxType = typeof(ISyntax);
            _grammar = Assembly.GetExecutingAssembly()
                               .GetTypes()
                               .Where(type => type.GetInterfaces().Contains(syntaxType))
                               .Cast<ISyntax>()
                               .ToList();
        }

        public SyntaxTreeNode Parse(CodeModule module)
        {
            var project = module.Parent.Collection.Parent;
            var component = module.Parent;

            var publicScope = project.Name;
            var localScope = string.Concat(project.Name, ".", component.Name);

            var content = SplitLogicalCodeLines(module.Lines[1, module.CountOfLines - module.CountOfDeclarationLines]);
            var memberNodes = ParseModuleMembers(publicScope, localScope, content);

            var result = new ModuleNode(project.Name, component.Name, memberNodes);

            return result;
        }

        private IDictionary<int, string> SplitLogicalCodeLines(string content)
        {
            const string lineContinuationMarker = "_";

            var result = new Dictionary<int, string>();
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
                    result.Add(index + 1, logicalLine.ToString());
                }
            }

            return result;
        }

        private IEnumerable<SyntaxTreeNode> ParseModuleMembers(string publicScope, string localScope, IDictionary<int, string> logicalCodeLines)
        {
            var currentLocalScope = localScope;
            var lines = logicalCodeLines.Values.ToList();
            for (int index = 0; index < lines.Count; index++)
            {
                var line = lines[index];
                var instructions = SplitInstructions(line);
                foreach (var instruction in instructions)
                {
                    SyntaxTreeNode node;
                    foreach (var syntax in _grammar)
                    {
                        node = syntax.ToNode(publicScope, currentLocalScope, instruction);
                        if (node != null)
                        {
                            if (node.DefinesScope)
                            {
                                
                            }
                            yield return node;
                        }
                    }
                }
            }
        }
        
        private IEnumerable<string> SplitInstructions(string logicalCodeLine)
        {
            const char instructionSeparator = ':';

            var trimmed = logicalCodeLine.Trim();
            if (!trimmed.Contains(' ') && trimmed.EndsWith(instructionSeparator.ToString()))
            {
                // line is a label; instruction separator also identifies a label...
                return new[] { logicalCodeLine };
            }

            return logicalCodeLine.Split(instructionSeparator)
                                  .Select(instruction => instruction.Trim());
        }
    }
}
