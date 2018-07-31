﻿using Rubberduck.VBEditor;

namespace Rubberduck.Common
{
    /// <summary>
    /// Represents a code string that includes caret position.
    /// </summary>
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

        public string[] Lines
        {
            get
            {
                return Code?.Split('\n') 
                    ?? new string[] { };
            }
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
            if (Code?.ToCodeString().CaretPosition != default)
            {
                return Code;
            }

            return InsertPseudoCaret();
        }

        private string InsertPseudoCaret()
        {
            var lines = (Code ?? string.Empty).Split('\n');
            var line = lines[CaretPosition.StartLine];
            lines[CaretPosition.StartLine] = line.Insert(CaretPosition.StartColumn, "|");
            return new CodeString(string.Join("\n", lines), CaretPosition);
        }
    }
}
