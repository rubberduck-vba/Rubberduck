using Rubberduck.VBEditor;
using System;
using System.Linq;

namespace Rubberduck.Common
{
    public class TestCodeString : CodeString
    {
        public static readonly char PseudoCaret = '|';

        public TestCodeString(CodeString codeString)
            : this(codeString.Code, codeString.CaretPosition, codeString.SnippetPosition)
        { }

        public TestCodeString(string code, Selection zPosition, Selection pPosition = default)
            : base(code, zPosition, pPosition)
        { }

        public override string ToString()
        {
            return InsertPseudoCaret();
        }

        private string InsertPseudoCaret()
        {
            if (string.IsNullOrEmpty(Code))
            {
                return string.Empty;
            }

            var lines = Lines;
            var line = lines[CaretPosition.StartLine];
            lines[CaretPosition.StartLine] = line.Insert(CaretPosition.StartColumn, PseudoCaret.ToString());
            return string.Join("\r\n", lines);
        }
    }

    /// <summary>
    /// Represents a code string that includes caret position.
    /// </summary>
    public class CodeString : IEquatable<CodeString>
    {
        /// <summary>
        /// Creates a new <c>CodeString</c>
        /// </summary>
        /// <param name="code">Code string</param>
        /// <param name="zPosition">Zero-based caret position in the code string.</param>
        /// <param name="pPosition">One-based selection span of the code string in the containing module.</param>
        public CodeString(string code, Selection zPosition, Selection pPosition = default)
        {
            Code = code ?? throw new ArgumentNullException(nameof(code));
            CaretPosition = zPosition;

            var lines = Lines;
            SnippetPosition = pPosition == default
                ? new Selection(1, 1, lines.Length, lines[lines.Length-1].Length)
                : pPosition;
        }

        /// <summary>
        /// The code string.
        /// </summary>
        public string Code { get; protected set; }
        /// <summary>
        /// Zero-based caret position in the code string.
        /// </summary>
        public Selection CaretPosition { get; }
        /// <summary>
        /// One-based position of the code string in the containing module.
        /// </summary>
        public Selection SnippetPosition { get; }
        /// <summary>
        /// Gets the individual <see cref="Code"/> string lines.
        /// </summary>
        public string[] Lines => Code?.Replace("\r", string.Empty).Split('\n') ?? new string[] { };

        public override bool Equals(object obj)
        {
            if (obj == null)
            {
                return false;
            }

            var other = (CodeString)obj;
            return Equals(other);
        }

        public override int GetHashCode()
        {
            return HashCode.Compute(Code, CaretPosition);
        }

        public override string ToString()
        {
            return Code;
        }

        public bool Equals(CodeString other)
        {
            if (other == null)
            {
                return false;
            }
            return (Code == null && other.Code == null) 
                || (Code != null && Code.Equals(other.Code) && CaretPosition.Equals(other.CaretPosition));
        }
    }
}
