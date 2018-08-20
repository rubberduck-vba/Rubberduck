using Rubberduck.VBEditor;
using System;

namespace Rubberduck.Common
{
    /// <summary>
    /// Represents a code string that includes caret position.
    /// </summary>
    public struct CodeString : IEquatable<CodeString>
    {
        /// <summary>
        /// Creates a new <c>CodeString</c>
        /// </summary>
        /// <param name="code">Code string</param>
        /// <param name="zPosition">Zero-based caret position in the code string.</param>
        /// <param name="pPosition">One-based selection span of the code string in the containing module.</param>
        public CodeString(string code, Selection zPosition, Selection pPosition = default)
        {
            if (code == null) throw new ArgumentNullException(nameof(code));

            var lines = code.Split('\n');
            var line = lines[zPosition.StartLine];
            if (line != string.Empty && line[Math.Min(line.Length - 1, zPosition.StartColumn)] == '|')
            {
                Code = line.Remove(Math.Min(line.Length - 1, zPosition.StartColumn), 1);
            }
            else
            {
                Code = code;
            }

            SnippetPosition = pPosition == default
                ? new Selection(1, 1, lines.Length, lines[lines.Length-1].Length)
                : pPosition;

            CaretPosition = zPosition;
        }

        public static CodeString FromString(string code)
        {
            var zPosition = new Selection();
            var lines = (code ?? string.Empty).Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                var line = lines[i];
                var index = line.IndexOf('|');
                if (index >= 0)
                {
                    lines[i] = line.Remove(index, 1);
                    zPosition = new Selection(i, index);
                    break;
                }
            }

            var newCode = string.Join("\n", lines);
            return new CodeString(code, zPosition);
        }


        /// <summary>
        /// The code string.
        /// </summary>
        public string Code { get; }
        /// <summary>
        /// Zero-based caret position in the code string.
        /// </summary>
        public Selection CaretPosition { get; }
        /// <summary>
        /// One-based position of the code string in the containing module.
        /// </summary>
        public Selection SnippetPosition { get; }

        public string[] Lines
        {
            get
            {
                return Code?.Split('\n') 
                    ?? new string[] { };
            }
        }

        public static bool operator ==(CodeString codeString1, CodeString codeString2) => (codeString1.Code == codeString2.Code && codeString1.CaretPosition == codeString2.CaretPosition);
        public static bool operator !=(CodeString codeString1, CodeString codeString2) => !(codeString1 == codeString2);

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
            return InsertPseudoCaret();
        }

        private string InsertPseudoCaret()
        {
            if (string.IsNullOrEmpty(Code))
            {
                return string.Empty;
            }

            var lines = Code.Split('\n');
            var line = lines[CaretPosition.StartLine];
            lines[CaretPosition.StartLine] = line.Insert(CaretPosition.StartColumn, "|");
            return string.Join("\n", lines);
        }

        public bool Equals(CodeString other)
        {
            if (other == default)
            {
                return false;
            }
            return (Code == null && other.Code == null) 
                || (Code != null && Code.Equals(other.Code) && CaretPosition.Equals(other.CaretPosition));
        }
    }
}
