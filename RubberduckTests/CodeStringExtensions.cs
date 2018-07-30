using System;
using Rubberduck.VBEditor;

namespace RubberduckTests
{
    public struct CodeString
    {
        /// <summary>
        /// Creates a new <c>CodeString</c>
        /// </summary>
        /// <param name="code">Code string</param>
        /// <param name="zPosition">Zero-based caret position</param>
        public CodeString(string code, Selection zPosition)
        {
            Code = code;
            CaretPosition = zPosition;
        }

        /// <summary>
        /// The code string.
        /// </summary>
        public string Code { get; }
        /// <summary>
        /// Zero-based caret position in the code string.
        /// </summary>
        public Selection CaretPosition { get; }

        public static implicit operator string(CodeString codeString)
        {
            return codeString.Code;
        }

        public override bool Equals(object obj)
        {
            if (obj == null)
            {
                return false;
            }

            var other = (CodeString)obj;
            return Code.Equals(other.Code) && CaretPosition.Equals(other.CaretPosition);
        }

        public override int GetHashCode()
        {
            return HashCode.Compute(Code, CaretPosition);
        }

        public override string ToString()
        {
            var hasCaret = Code.RemovePseudoCaret().CaretPosition != default;
            if (hasCaret)
            {
                return Code;
            }

            return Code.InsertPseudoCaret(CaretPosition);
        }
    }

    public static class CodeStringExtensions
    {
        public static CodeString RemovePseudoCaret(this string code)
        {
            var zPosition = new Selection();
            var lines = code.Split('\n');
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
            return new CodeString(newCode, zPosition);
        }

        public static CodeString InsertPseudoCaret(this string code, Selection zPosition)
        {
            var lines = code.Split('\n');
            var line = lines[zPosition.StartLine];
            lines[zPosition.StartLine] = line.Insert(zPosition.StartColumn, "|");
            return new CodeString(string.Join("\n", lines), zPosition);
        }
    }
}
