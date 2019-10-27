using System.Collections.Generic;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Inspections.Inspections.Extensions;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Identifies the use of indexed default member accesses for which the default member cannot be determined at compile time.
    /// </summary>
    /// <why>
    /// An indexed default member access hides away the actually called member. This is especially problematic if the default member cannot be determined from the declared type of the object.
    /// Should there not be a suitable default member at runtime, an error 438 'Object doesn't support this property or method' will be raised.
    /// </why>
    /// <example hasresult="true">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal rst As Object)
    ///     Dim bar As Variant
    ///     bar = rst("MyField")
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasresult="false">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal rst As Object)
    ///     Dim bar As Variant
    ///     bar = rst.Fields.Item("MyField")
    /// End Sub
    /// ]]>
    /// </example>
    public sealed class IndexedUnboundDefaultMemberAccessInspection : IdentifierReferenceInspectionBase
    {
        public IndexedUnboundDefaultMemberAccessInspection(RubberduckParserState state)
            : base(state)
        {
            Severity = CodeInspectionSeverity.Warning;
        }

        protected override IEnumerable<IdentifierReference> ReferencesInModule(QualifiedModuleName module)
        {
            return DeclarationFinderProvider.DeclarationFinder.UnboundDefaultMemberAccesses(module);
        }

        protected override bool IsResultReference(IdentifierReference reference)
        {
            return reference.IsIndexedDefaultMemberAccess
                   && !(reference.Context is VBAParser.DictionaryAccessContext)
                   && !reference.IsIgnoringInspectionResultFor(AnnotationName);
        }

        protected override string ResultDescription(IdentifierReference reference)
        {
            var expression = reference.IdentifierName;
            return string.Format(InspectionResults.IndexedUnboundDefaultMemberAccessInspection, expression);
        }
    }
}