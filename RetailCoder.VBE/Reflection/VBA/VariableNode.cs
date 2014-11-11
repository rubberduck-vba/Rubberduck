using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Rubberduck.Reflection.VBA
{
    // todo: handle multiple declarations on single instruction - grammar/regex already supports it.

    internal class VariableNode : DeclarationNodeBase
    {
        public VariableNode(string scope, Match match, string comment)
            : base(scope, match, comment)
        { }

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
