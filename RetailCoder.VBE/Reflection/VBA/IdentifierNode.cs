using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace Rubberduck.Reflection.VBA
{
    [ComVisible(false)]
    public class IdentifierNode : SyntaxTreeNode
    {
        public static readonly IDictionary<string, string> TypeSpecifiers = new Dictionary<string, string>
        {
            { "%", ReservedKeywords.Integer },
            { "&", ReservedKeywords.Long },
            { "@", ReservedKeywords.Decimal },
            { "!", ReservedKeywords.Single },
            { "#", ReservedKeywords.Double },
            { "$", ReservedKeywords.String }
        };

        public IdentifierNode(Instruction instruction, string scope, Match match)
            : base(instruction, scope, match)
        {
        }

        public virtual string Identifier
        {
            get
            {
                return RegexMatch.Groups["identifier"].Success
                    ? RegexMatch.Groups["identifier"].Value
                    : string.Empty;
            }
        }

        /// <summary>
        /// Gets the type specifier character, if the declaration included one.
        /// </summary>
        /// <remarks>
        /// <see cref="TypeSpecifier"/> and <see cref="Reference"/> are mutually exclusive.
        /// </remarks>
        public virtual string TypeSpecifier
        {
            get
            {
                return RegexMatch.Groups["specifier"].Success
                    ? RegexMatch.Groups["specifier"].Value
                    : string.Empty;
            }
        }

        /// <summary>
        /// Gets the qualified type name the identifier is declared with.
        /// </summary>
        /// <example>
        /// Returns "ADODB.Recordset" from <c>Dim rs As ADODB.Recordset</c> declaration.
        /// </example>
        /// <remarks>
        /// <see cref="TypeSpecifier"/> and <see cref="Reference"/> are mutually exclusive.
        /// </remarks>
        public virtual string Reference
        {
            get
            {
                return RegexMatch.Groups["reference"].Success
                    ? RegexMatch.Groups["reference"].Value
                    : string.Empty;
            }
        }

        /// <summary>
        /// Gets the type name of the identifier, regardless of how it is declared.
        /// </summary>
        public virtual string TypeName
        {
            get
            {
                return IsTypeSpecified
                    ? !string.IsNullOrEmpty(TypeSpecifier)
                        ? TypeSpecifiers[TypeSpecifier]
                        : Reference
                    : ReservedKeywords.Variant;
            }
        }

        /// <summary>
        /// Gets a value indicating whether the type was specified in the declaration.
        /// </summary>
        /// <remarks>
        /// Useful for telling implicit from explicit <c>Variant</c> declarations.
        /// </remarks>
        /// <example>
        /// Returns <c>true</c> with <c>Dim foo As Variant</c> and <c>Dim foo$</c> declarations, and <c>false</c> with <c>Dim foo</c> declaration.
        /// </example>
        public virtual bool IsTypeSpecified
        {
            get
            {
                return !string.IsNullOrEmpty(TypeSpecifier) || RegexMatch.Groups["as"].Success;
            }
        }

        /// <summary>
        /// If declaration specifies it, gets the library the type is qualified with.
        /// </summary>
        /// <example>
        /// Returns "ADODB" from <c>Dim rs As ADODB.Recordset</c> declaration.
        /// </example>
        public virtual string Library
        {
            get
            {
                return RegexMatch.Groups["library"].Success
                    ? RegexMatch.Groups["library"].Value
                    : string.Empty;
            }
        }


        /// <summary>
        /// Gets a value indicating whether declaration is initializing variable with a <c>New</c> instance of a class.
        /// </summary>
        /// <example>
        /// Returns <c>true</c> with <c>Dim foo As New Bar</c> declaration.
        /// </example>
        public bool IsInitialized { get { return RegexMatch.Groups["initializer"].Success; } }

        /// <summary>
        /// Gets a value indicating whether declaration is an array.
        /// </summary>
        /// <exexample>
        /// Returns <c>true</c> with <c>Dim foo() As String</c>, <c>Dim foo(10) As String</c> and <c>Dim foo(1 To 10) As String</c> declarations.
        /// </exexample>
        public bool IsArray { get { return RegexMatch.Groups["array"].Success; } }

        /// <summary>
        /// Gets the number of dimensions in an array declaration.
        /// Returns 0 if declaration is not an array, or if array is dynamically sized.
        /// </summary>
        /// <example>
        /// Returns <c>2</c> with <c>Dim foo(10, 2)</c>
        /// </example>
        public int ArrayDimensionsCount
        {
            get
            {
                if (IsArray && RegexMatch.Groups["size"].Success)
                {
                    return RegexMatch.Groups["size"].Value.Count(c => c == ',') + 1;
                }
                else
                {
                    return 0;
                }
            }
        }

        /// <summary>
        /// Gets a value indicating whether declaration is a dynamically sized array.
        /// Returns <c>false</c> if declaration is not an array.
        /// </summary>
        /// <example>
        /// Returns <c>true</c> with <c>Dim foo() As String</c> declaration.
        /// </example>
        public bool IsDynamicArray { get { return IsArray && !RegexMatch.Groups["size"].Success; } }
    }
}