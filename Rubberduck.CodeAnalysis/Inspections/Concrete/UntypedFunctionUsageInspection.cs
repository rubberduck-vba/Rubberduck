using System.Collections.Generic;
using System.Linq;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Flags uses of a number of specific string-centric but Variant-returning functions in various standard library modules.
    /// </summary>
    /// <why>
    /// Several functions in the standard library take a Variant parameter and return a Variant result, but an equivalent 
    /// string-returning function taking a string parameter exists and should probably be preferred.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal foo As Double)
    ///     Debug.Print Format(foo, "Currency") ' Strings.Format function returns a Variant.
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal foo As Double)
    ///     Debug.Print Format$(CStr(foo), "Currency") ' Strings.Format$ function returns a String.
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class UntypedFunctionUsageInspection : IdentifierReferenceInspectionFromDeclarationsBase
    {
        public UntypedFunctionUsageInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {}

        private readonly HashSet<string> _tokens = new HashSet<string>{
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

        protected override IEnumerable<Declaration> ObjectionableDeclarations(DeclarationFinder finder)
        {
            return BuiltInVariantStringFunctionsWithStringTypedVersion(finder);
        }

        private IEnumerable<Declaration> BuiltInVariantStringFunctionsWithStringTypedVersion(DeclarationFinder finder)
        {
            return finder
                .BuiltInDeclarations(DeclarationType.Member)
                .Where(item => item.Scope.StartsWith("VBE7.DLL;") 
                               && (_tokens.Contains(item.IdentifierName)
                                    || item.IdentifierName.StartsWith("_B_var_")
                                        && _tokens.Contains(item.IdentifierName.Substring("_B_var_".Length))));
        }

        protected override string ResultDescription(IdentifierReference reference)
        {
            var declarationName = reference.Declaration.IdentifierName;
            return string.Format(
                InspectionResults.UntypedFunctionUsageInspection,
                declarationName);
        }

        protected override bool IsResultReference(IdentifierReference reference, DeclarationFinder finder)
        {
            return IsNotStringHinted(reference);
        }

        private bool IsNotStringHinted(IdentifierReference reference)
        {
            return _tokens.Contains(reference.IdentifierName);
        }
    }
}
