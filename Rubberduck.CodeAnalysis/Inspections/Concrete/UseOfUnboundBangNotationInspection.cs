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
    /// Identifies the use of bang notation, formally known as dictionary access expression, for which the default member is not known at compile time.
    /// </summary>
    /// <why>
    /// A dictionary access expression looks like a strongly typed call, but it actually is a stringly typed access to the parameterized default member of the object.
    /// This is especially misleading the default member cannot be determined at compile time.  
    /// </why>
    /// <example hasresult="true">
    /// <![CDATA[
    /// Public Function MyName(ByVal rst As Object) As Variant
    ///     MyName = rst!Name.Value
    /// End Function
    /// ]]>
    /// </example>
    /// <example hasresult="true">
    /// <![CDATA[
    /// Public Function MyName(ByVal rst As Variant) As Variant
    ///     With rst
    ///         MyName = !Name.Value
    ///     End With
    /// End Function
    /// ]]>
    /// </example>
    /// <example hasresult="false">
    /// <![CDATA[
    /// Public Function MyName(ByVal rst As ADODB.Recordset) As Variant
    ///     MyName = rst!Name.Value
    /// End Function
    /// ]]>
    /// </example>
    /// <example hasresult="false">
    /// <![CDATA[
    /// Public Function MyName(ByVal rst As Object) As Variant
    ///     MyName = rst("Name").Value
    /// End Function
    /// ]]>
    /// </example>
    /// <example hasresult="false">
    /// <![CDATA[
    /// Public Function MyName(ByVal rst As Variant) As Variant
    ///     With rst
    ///         MyName = .Fields.Item("Name").Value
    ///     End With
    /// End Function
    /// ]]>
    /// </example>
    public sealed class UseOfUnboundBangNotationInspection : IdentifierReferenceInspectionBase
    {
        public UseOfUnboundBangNotationInspection(RubberduckParserState state)
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
                   && reference.Context is VBAParser.DictionaryAccessContext;
        }

        protected override string ResultDescription(IdentifierReference reference)
        {
            var expression = reference.IdentifierName;
            return string.Format(InspectionResults.UseOfRecursiveBangNotationInspection, expression);
        }
    }
}