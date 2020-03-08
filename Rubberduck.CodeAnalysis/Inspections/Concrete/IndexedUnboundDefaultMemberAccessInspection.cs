using System.Collections.Generic;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.CodeAnalysis.Inspections.Extensions;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;
using Rubberduck.VBEditor;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Identifies the use of indexed default member accesses for which the default member cannot be determined at compile time.
    /// </summary>
    /// <why>
    /// An indexed default member access hides away the actually called member. This is especially problematic if the default member cannot be determined from the declared type of the object.
    /// Should there not be a suitable default member at runtime, an error 438 'Object doesn't support this property or method' will be raised.
    /// </why>
    /// <example hasresult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal rst As Object)
    ///     Dim bar As Variant
    ///     bar = rst("MyField")
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasresult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal rst As Object)
    ///     Dim bar As Variant
    ///     bar = rst.Fields.Item("MyField")
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class IndexedUnboundDefaultMemberAccessInspection : IdentifierReferenceInspectionBase
    {
        public IndexedUnboundDefaultMemberAccessInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {
            Severity = CodeInspectionSeverity.Warning;
        }

        protected override IEnumerable<IdentifierReference> ReferencesInModule(QualifiedModuleName module, DeclarationFinder finder)
        {
            return finder.UnboundDefaultMemberAccesses(module);
        }

        protected override bool IsResultReference(IdentifierReference reference, DeclarationFinder finder)
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