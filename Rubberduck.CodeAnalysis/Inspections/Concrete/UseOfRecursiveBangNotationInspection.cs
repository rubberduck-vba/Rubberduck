using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Identifies the use of bang notation, formally known as dictionary access expression, for which a recursive default member resolution is necessary.
    /// </summary>
    /// <why>
    /// A dictionary access expression looks like a strongly typed call, but it actually is a stringly typed access to the parameterized default member of the object.
    /// This is especially misleading if the parameterized default member is not on the object itself and can only be reached by calling the parameterless default member first.  
    /// </why>
    /// <example hasresult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Function MyName(ByVal rst As ADODB.Recordset) As Variant
    ///     MyName = rst!Name.Value
    /// End Function
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasresult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Function MyName(ByVal rst As ADODB.Recordset) As Variant
    ///     With rst
    ///         MyName = !Name.Value
    ///     End With
    /// End Function
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasresult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Function MyName(ByVal rst As ADODB.Recordset) As Variant
    ///     MyName = rst.Fields.Item("Name").Value
    /// End Function
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasresult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Function MyName(ByVal rst As ADODB.Recordset) As Variant
    ///     MyName = rst("Name").Value
    /// End Function
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasresult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Function MyName(ByVal rst As ADODB.Recordset) As Variant
    ///     MyName = rst.Fields!Name.Value 'see "UseOfBangNotation" inspection
    /// End Function
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasresult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Function MyName(ByVal rst As ADODB.Recordset) As Variant
    ///     With rst
    ///         MyName = .Fields.Item("Name").Value
    ///     End With
    /// End Function
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class UseOfRecursiveBangNotationInspection : IdentifierReferenceInspectionBase
    {
        public UseOfRecursiveBangNotationInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {
            Severity = CodeInspectionSeverity.Suggestion;
        }

        protected override bool IsResultReference(IdentifierReference reference, DeclarationFinder finder)
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