using System.Runtime.InteropServices;
using Rubberduck.Extensions;
using Rubberduck.VBA.Parser.Grammar;

namespace Rubberduck.VBA.Parser
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
            _startColumn = startColumn;
            _endColumn = endColumn;
            _content = content;

            int index;
            if (_content.HasComment(out index))
            {
                _comment = _content.Substring(index);
                _instruction = _content.Substring(0, index);
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