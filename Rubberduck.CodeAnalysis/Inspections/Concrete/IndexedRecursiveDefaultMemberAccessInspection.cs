using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.CodeAnalysis.Inspections.Extensions;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Identifies the use of indexed default member accesses that require a recursive default member resolution.
    /// </summary>
    /// <why>
    /// An indexed default member access hides away the actually called member. This is especially problematic if the corresponding parameterized default member is not on the interface of the object itself.
    /// </why>
    /// <example hasresult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal rst As ADODB.Recordset)
    ///     Dim bar As Variant
    ///     bar = rst("MyField")
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasresult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal rst As ADODB.Recordset)
    ///     Dim bar As Variant
    ///     bar = rst.Fields.Item("MyField")
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class IndexedRecursiveDefaultMemberAccessInspection : IdentifierReferenceInspectionBase
    {
        public IndexedRecursiveDefaultMemberAccessInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {
            Severity = CodeInspectionSeverity.Suggestion;
        }

        protected override bool IsResultReference(IdentifierReference reference, DeclarationFinder finder)
        {
            return reference.IsIndexedDefaultMemberAccess
                   && reference.DefaultMemberRecursionDepth > 1
                   && !(reference.Context is VBAParser.DictionaryAccessContext)
                   && !reference.IsIgnoringInspectionResultFor(AnnotationName);
        }

        protected override string ResultDescription(IdentifierReference reference)
        {
            var expression = reference.IdentifierName;
            var defaultMember = reference.Declaration.QualifiedName.ToString();
            return string.Format(InspectionResults.IndexedRecursiveDefaultMemberAccessInspection, expression, defaultMember);
        }
    }
}