using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Rubberduck.VBA.Parser.Grammar;

namespace Rubberduck.VBA.Parser
{
    /// <summary>
    /// An immutable value type representing a line of code. Abstracts line continuations.
    /// </summary>
    [ComVisible(false)]
    public struct LogicalCodeLine
    {
        public LogicalCodeLine(string projectName, string componentName, int startLineNumber, int endLineNumber, string content)
        {
            _projectName = projectName;
            _componentName = componentName;
            _startLineNumber = startLineNumber;
            _endLineNumber = endLineNumber;
            _content = content;
        }

        private readonly string _projectName;
        /// <summary>
        /// The name of the project this logical code line belongs to.
        /// </summary>
        public string ProjectName { get { return _projectName; } }

        private readonly string _componentName;
        /// <summary>
        /// The name of the project component this logical code line belongs to.
        /// </summary>
        public string ComponentName { get { return _componentName; } }

        private readonly int _endLineNumber;
        /// <summary>
        /// The code pane line number where this logical code line ends.
        /// </summary>
        public int EndLineNumber { get { return _endLineNumber; } }

        private readonly int _startLineNumber;
        /// <summary>
        /// The code pane line number where this logical code line begins.
        /// </summary>
        public int StartLineNumber { get { return _startLineNumber; } }

        private readonly string _content;
        /// <summary>
        /// The integral content of the logical code line, including line continuations.
        /// </summary>
        public string Content { get { return _content; } }

        /// <summary>
        /// Splits a logical code line into a number of <see cref="Instruction"/> instances.
        /// </summary>
        /// <returns></returns>
        public IEnumerable<Instruction> SplitInstructions()
        {
            // return empty instruction for empty line to preserve vertical whitespace:
            if (string.IsNullOrWhiteSpace(_content))
            {
                return new[] { Instruction.Empty(this) };
            }

            const char separator = ':';

            // LabelSyntax uses instruction separator; 
            // return entire line if there's no separator or if LabelSyntax matches:
            var stripped = _content.StripTrailingComment();
            if (!stripped.Contains(separator) || Regex.Match(stripped, VBAGrammar.LabelSyntax).Success)
            {
                var indentation = stripped.TakeWhile(char.IsWhiteSpace).Count() + 1;
                return new[] { new Instruction(this, indentation, stripped.Length, _content) };
            }

            var result = new List<Instruction>();
            var instructionsCount = stripped.Count(c => c == separator) + 1;
            var startIndex = 0;
            var endIndex = 0;
            for (var instruction = 0; instruction < instructionsCount; instruction++)
            {
                endIndex = instruction == instructionsCount - 1
                    ? _content.Length
                    : _content.IndexOf(separator, endIndex);

                result.Add(new Instruction(this, startIndex, endIndex + 1, _content.Substring(startIndex, endIndex - startIndex)));
                startIndex = endIndex;
            }

            return result;
        }
    }
}
