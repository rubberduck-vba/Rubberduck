using Rubberduck.VBEditor;

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
            var hasCaret = Code.ToCodeString().CaretPosition != default;
            if (hasCaret)
            {
                return Code;
            }

            return Code.InsertPseudoCaret(CaretPosition);
        }
    }
}
