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
    /// Identifies the use of indexed default member accesses.
    /// </summary>
    /// <why>
    /// An indexed default member access hides away the actually called member.
    /// </why>
    /// <example hasresult="true">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal coll As Collection)
    ///     Dim bar As Variant
    ///     bar = coll(23)
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasresult="false">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal coll As Collection)
    ///     Dim bar As Variant
    ///     bar = coll.Item(23)
    /// End Sub
    /// ]]>
    /// </example>
    public sealed class IndexedDefaultMemberAccessInspection : IdentifierReferenceInspectionBase
    {
        public IndexedDefaultMemberAccessInspection(RubberduckParserState state)
            : base(state)
        {
            Severity = CodeInspectionSeverity.Hint;
        }

        protected override bool IsResultReference(IdentifierReference reference)
        {
            return reference.IsIndexedDefaultMemberAccess
                   && reference.DefaultMemberRecursionDepth == 1
                   && !(reference.Context is VBAParser.DictionaryAccessContext)
                   && !reference.IsIgnoringInspectionResultFor(AnnotationName);
        }

        protected override string ResultDescription(IdentifierReference reference)
        {
            var expression = reference.IdentifierName;
            var defaultMember = reference.Declaration.QualifiedName.ToString();
            return string.Format(InspectionResults.IndexedDefaultMemberAccessInspection, expression, defaultMember);
        }
    }
}