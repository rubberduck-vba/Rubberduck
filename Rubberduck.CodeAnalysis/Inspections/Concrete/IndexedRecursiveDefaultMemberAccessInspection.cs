using Rubberduck.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Inspections.Inspections.Extensions;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Identifies the use of indexed default member accesses that require a recursive default member resolution.
    /// </summary>
    /// <why>
    /// An indexed default member access hides away the actually called member. This is especially problematic if the corresponding parameterized default member is not on the interface of the object itself.
    /// </why>
    /// <example hasresult="true">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal rst As ADODB.Recordset)
    ///     Dim bar As Variant
    ///     bar = rst("MyField")
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasresult="false">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal rst As ADODB.Recordset)
    ///     Dim bar As Variant
    ///     bar = rst.Fields.Item("MyField")
    /// End Sub
    /// ]]>
    /// </example>
    public sealed class IndexedRecursiveDefaultMemberAccessInspection : IdentifierReferenceInspectionBase
    {
        public IndexedRecursiveDefaultMemberAccessInspection(RubberduckParserState state)
            : base(state)
        {
            Severity = CodeInspectionSeverity.Suggestion;
        }

        protected override bool IsResultReference(IdentifierReference reference)
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