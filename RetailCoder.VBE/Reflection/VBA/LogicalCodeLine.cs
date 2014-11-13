using Rubberduck.Reflection.VBA.Grammar;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Rubberduck.Reflection.VBA
{
    /// <summary>
    /// An immutable value type representing a line of code. Abstracts line continuations.
    /// </summary>
    internal struct LogicalCodeLine
    {
        public LogicalCodeLine(string projectName, string componentName, int lineNumber, string content)
        {
            _projectName = projectName;
            _componentName = componentName;
            _lineNumber = lineNumber;
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

        private readonly int _lineNumber;
        /// <summary>
        /// The code pane line number where this logical code line begins.
        /// </summary>
        public int LineNumber { get { return _lineNumber; } }

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
            if (!_content.Contains(separator) || Regex.Match(_content.StripTrailingComment(), VBAGrammar.LabelSyntax()).Success)
            {
                return new[] { new Instruction(this, 1, _content.Length, _content) };
            }

            var result = new List<Instruction>();
            var instructionsCount = _content.Count(c => c == separator) + 1;
            var startIndex = 0;
            var endIndex = 0;
            for (int instruction = 0; instruction < instructionsCount; instruction++)
            {
                endIndex = instruction == instructionsCount - 1 
                    ? _content.Length
                    : _content.IndexOf(separator, endIndex);

                result.Add(new Instruction(this, startIndex, endIndex, _content.Substring(startIndex, endIndex - startIndex)));
                startIndex = endIndex;
            }

            return result;
        }
    }

    /// <summary>
    /// An immutable value type representing a single instruction. Abstracts instruction separators.
    /// </summary>
    internal struct Instruction
    {
        public Instruction(LogicalCodeLine line, int startColumn, int endColumn, string content)
        {
            _line = line;
            _startColumn = startColumn;
            _endColumn = endColumn;
            _content = content;

            int index;
            if (_content.HasComment(out index))
            {
                _comment = _content.Substring(index);
                _instruction = _content.Substring(0, index - 1);
            }
            else
            {
                _comment = string.Empty;
                _instruction = _content;
            }
        }

        private readonly LogicalCodeLine _line;
        /// <summary>
        /// The <see cref="LogicalCodeLine"/> that contains this instruction.
        /// </summary>
        public LogicalCodeLine Line { get { return _line; } }

        private readonly int _startColumn;
        /// <summary>
        /// The code pane column where this instruction begins.
        /// </summary>
        public int StartColumn { get { return _startColumn; } }

        private readonly int _endColumn;
        /// <summary>
        /// The code pane column where this instruction ends.
        /// </summary>
        public int EndColumn { get { return _endColumn; } }

        private readonly string _content;
        /// <summary>
        /// The entire instruction, including any trailing comment.
        /// </summary>
        public string Content { get { return _content; } }

        private readonly string _instruction;
        /// <summary>
        /// The instruction string, without any trailing comment.
        /// </summary>
        public string Value { get { return _instruction; } }

        private readonly string _comment;
        /// <summary>
        /// The trailing comment, if any.
        /// </summary>
        public string Comment { get { return _comment; } }

        public static Instruction Empty(LogicalCodeLine line)
        {
            return new Instruction(line, 1, 1, string.Empty);
        }
    }
}
