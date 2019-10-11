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
    /// Identifies the use of bang notation, formally known as dictionary access expression, for which a recursive default member resolution is necessary.
    /// </summary>
    /// <why>
    /// A dictionary access expression looks like a strongly typed call, but it actually is a stringly typed access to the parameterized default member of the object.
    /// This is especially misleading if the parameterized default member is not on the object itself and can only be reached by calling the parameterless default member first.  
    /// </why>
    /// <example hasresult="true">
    /// <![CDATA[
    /// Public Function MyName(ByVal rst As ADODB.Recordset) As Variant
    ///     MyName = rst!Name.Value
    /// End Function
    /// ]]>
    /// </example>
    /// <example hasresult="true">
    /// <![CDATA[
    /// Public Function MyName(ByVal rst As ADODB.Recordset) As Variant
    ///     With rst
    ///         MyName = !Name.Value
    ///     End With
    /// End Function
    /// ]]>
    /// </example>
    /// <example hasresult="false">
    /// <![CDATA[
    /// Public Function MyName(ByVal rst As ADODB.Recordset) As Variant
    ///     MyName = rst.Fields.Item("Name").Value
    /// End Function
    /// ]]>
    /// </example>
    /// <example hasresult="false">
    /// <![CDATA[
    /// Public Function MyName(ByVal rst As ADODB.Recordset) As Variant
    ///     MyName = rst("Name").Value
    /// End Function
    /// ]]>
    /// </example>
    /// <example hasresult="false">
    /// <![CDATA[
    /// Public Function MyName(ByVal rst As ADODB.Recordset) As Variant
    ///     MyName = rst.Fields!Name.Value 'see "UseOfBangNotation" inspection
    /// End Function
    /// ]]>
    /// </example>
    /// <example hasresult="false">
    /// <![CDATA[
    /// Public Function MyName(ByVal rst As ADODB.Recordset) As Variant
    ///     With rst
    ///         MyName = .Fields.Item("Name").Value
    ///     End With
    /// End Function
    /// ]]>
    /// </example>
    public sealed class UseOfRecursiveBangNotationInspection : IdentifierReferenceInspectionBase
    {
        public UseOfRecursiveBangNotationInspection(RubberduckParserState state)
            : base(state)
        {
            Severity = CodeInspectionSeverity.Suggestion;
        }

        protected override bool IsResultReference(IdentifierReference reference)
        {
            return reference.IsIndexedDefaultMemberAccess
                   && reference.DefaultMemberRecursionDepth > 1
                   && reference.Context is VBAParser.DictionaryAccessContext;
        }

        protected override string ResultDescription(IdentifierReference reference)
        {
            var expression = reference.IdentifierName;
            return string.Format(InspectionResults.UseOfRecursiveBangNotationInspection, expression);
        }
    }
}