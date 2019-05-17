using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.VBA;
using Rubberduck.Inspections.Inspections.Extensions;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Flags uses of a number of specific string-centric but Variant-returning functions in various standard library modules.
    /// </summary>
    /// <why>
    /// Several functions in the standard library take a Variant parameter and return a Variant result, but an equivalent 
    /// string-returning function taking a string parameter exists and should probably be preferred.
    /// </why>
    /// <example>
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal foo As Double)
    ///     Debug.Print Format(foo, "Currency") ' Strings.Format function returns a Variant.
    /// End Sub
    /// ]]>
    /// </example>
    /// <example>
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal foo As Double)
    ///     Debug.Print Format$(CStr(foo), "Currency") ' Strings.Format$ function returns a String.
    /// End Sub
    /// ]]>
    /// </example>
    public sealed class UntypedFunctionUsageInspection : InspectionBase
    {
        public UntypedFunctionUsageInspection(RubberduckParserState state)
            : base(state) { }

        private readonly string[] _tokens = {
            Tokens.Error,
            Tokens.Hex,
            Tokens.Oct,
            Tokens.Str,
            Tokens.CurDir,
            Tokens.Command,
            Tokens.Environ,
            Tokens.Chr,
            Tokens.ChrW,
            Tokens.Format,
            Tokens.Input,
            Tokens.InputB,
            Tokens.LCase,
            Tokens.Left,
            Tokens.LeftB,
            Tokens.LTrim,
            Tokens.Mid,
            Tokens.MidB,
            Tokens.Trim,
            Tokens.Right,
            Tokens.RightB,
            Tokens.RTrim,
            Tokens.UCase
        };

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var declarations = BuiltInDeclarations
                .Where(item =>
                        _tokens.Any(token => item.IdentifierName == token || item.IdentifierName == "_B_var_" + token) &&
                        item.Scope.StartsWith("VBE7.DLL;"));

            return declarations.SelectMany(declaration => declaration.References
                .Where(item => _tokens.Contains(item.IdentifierName) &&
                               !item.IsIgnoringInspectionResultFor(AnnotationName))
                .Select(item => new IdentifierReferenceInspectionResult(this,
                                                     string.Format(InspectionResults.UntypedFunctionUsageInspection, item.Declaration.IdentifierName),
                                                     State,
                                                     item)));
        }
    }
}
