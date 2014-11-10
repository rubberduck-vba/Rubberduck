using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace RetailCoderVBE.Reflection.VBA
{
    // todo: handle multiple declarations on single instruction - grammar/regex already supports it.

    /// <summary>
    /// Base class for a declaration node, in the form of <c>{Keyword} {Identifier}[TypeSpecifier] [As [New ]{Reference}]</c>.
    /// </summary>
    internal abstract class DeclarationNode : SyntaxTreeNode
    {
        public DeclarationNode(string scope, Match match, string comment)
            : base(scope, match, comment)
        { }

        private static IDictionary<string, string> _typeSpecifiers = new Dictionary<string, string>
            {
                { "%", ReservedKeywords.Integer },
                { "&", ReservedKeywords.Long },
                { "@", ReservedKeywords.Decimal },
                { "!", ReservedKeywords.Single },
                { "#", ReservedKeywords.Double },
                { "$", ReservedKeywords.String }
            };

        /// <summary>
        /// Gets the declared identifier.
        /// </summary>
        /// <example>
        /// Returns "foo" from <c>Dim foo As String</c> declaration.
        /// </example>
        public Identifier Identifier
        {
            get
            {
                var name = RegexMatch.Groups["identifier"].Value;
                return new Identifier(Scope, name, TypeName);
            }
        }

        /// <summary>
        /// Gets the type specifier character, if the declaration included one.
        /// </summary>
        /// <remarks>
        /// <see cref="TypeSpecifier"/> and <see cref="Reference"/> are mutually exclusive.
        /// </remarks>
        public string TypeSpecifier
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
        public string Reference
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
        public string TypeName
        {
            get
            {
                return IsTypeSpecified
                    ? !string.IsNullOrEmpty(TypeSpecifier)
                        ? _typeSpecifiers[TypeSpecifier]
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
        public bool IsTypeSpecified
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
        public string Library
        {
            get
            {
                return RegexMatch.Groups["library"].Success
                    ? RegexMatch.Groups["library"].Value
                    : string.Empty;
            }
        }
    }
}
