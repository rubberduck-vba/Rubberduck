﻿using System;
using System.Diagnostics;

namespace Rubberduck.VBEditor
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
            lines[CaretPosition.StartLine] = line.Insert(Math.Min(CaretPosition.StartColumn, line.Length), PseudoCaret.ToString());
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
        public string Code { get; }
        /// <summary>
        /// Zero-based caret position in the code string.
        /// </summary>
        public Selection CaretPosition { get; }
        /// <summary>
        /// Gets the 0-based index of the caret position in the flattened <see cref="Code"/> string.
        /// </summary>
        public int CaretCharIndex
        {
            get
            {
                var i = 0;
                for (var line = 0; line <= CaretPosition.StartLine; line++)
                {
                    if (line < CaretPosition.StartLine)
                    {
                        i += Lines[line].Length;
                    }
                    else
                    {
                        i += CaretPosition.StartColumn;
                        return i;
                    }

                    i += 2; // "\r\n"
                }

                return i;
            }
        }
        /// <summary>
        /// One-based position of the code string in the containing module.
        /// </summary>
        public Selection SnippetPosition { get; }
        /// <summary>
        /// Gets the individual <see cref="Code"/> string lines.
        /// </summary>
        public string[] Lines => Code?.Replace("\r", string.Empty).Split('\n') ?? new string[] { };
        /// <summary>
        /// Gets the contents of the line that is immediately before the line that contains the caret.
        /// </summary>
        public string PreviousLine => CaretPosition.StartLine == 0 ? null : Lines[CaretPosition.StartLine - 1];
        /// <summary>
        /// Gets the contents of the line that is immediately after the line that contains the caret.
        /// </summary>
        public string NextLine => CaretPosition.StartLine == Lines.Length ? null : Lines[CaretPosition.StartLine + 1];

        /// <summary>
        /// Gets the contents of the line that contains the caret.
        /// </summary>
        public string CaretLine => Lines[CaretPosition.StartLine];

        public CodeString ReplaceLine(int index, string content)
        {
            var lines = Lines;
            Debug.Assert(index >= 0 && index < lines.Length);

            lines[index] = content;
            var code = string.Join("\r\n", lines);
            return new CodeString(code, CaretPosition, SnippetPosition);
        }

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
