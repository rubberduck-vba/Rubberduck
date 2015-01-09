using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.Extensions;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.VBA
{
    /// <summary>
    /// An immutable value type representing a single instruction. Abstracts instruction separators.
    /// </summary>
    [ComVisible(false)]
    public struct Instruction
    {
        public Instruction(LogicalCodeLine line, int startColumn, int endColumn, string content)
        {
            _line = line;
            _startColumn = startColumn == 0 ? 1 : startColumn;
            _endColumn = endColumn == 0 ? startColumn : endColumn;
            _content = content;

            int index;
            if (_content.HasComment(out index))
            {
                _comment = _content.TrimStart().Substring(index - _content.TakeWhile(c => c == ' ').Count()).Trim();
                _instruction = _content.TrimStart().Substring(0, index - _content.TakeWhile(c => c == ' ').Count());
            }
            else
            {
                _comment = string.Empty;
                _instruction = _content.TrimEnd().EndsWith(":") ? _content.TrimEnd().Substring(0, _content.TrimEnd().Length - 1) : _content;
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

        public Selection Selection 
        {
            get { return new Selection(_line.StartLineNumber, _startColumn, _line.EndLineNumber, _endColumn); }
        }

        public static Instruction Empty(LogicalCodeLine line)
        {
            return new Instruction(line, 1, 1, string.Empty);
        }
    }
}