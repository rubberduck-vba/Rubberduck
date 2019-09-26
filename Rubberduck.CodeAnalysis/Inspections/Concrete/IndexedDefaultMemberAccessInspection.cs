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
    /// Identifies the use indexed default member accesses.
    /// </summary>
    /// <why>
    /// An indexed default member access hides away the actually called member.
    /// </why>
    /// <example hasResult="true">
    /// <![CDATA[
    /// Class1:
    ///
    /// Public Function Foo(ByVal arg As Long) As Long
    /// Attibute VB_UserMemId = 0
    ///     Foo = 42
    /// End Function
    ///
    /// Module:
    /// 
    /// Public Sub DoSomething(ByVal arg As Class1)
    ///     Dim bar As Variant
    ///     bar = arg(23)
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResult="false">
    /// <![CDATA[
    /// Class1:
    ///
    /// Public Function Foo(ByVal arg As Long) As Long
    /// Attibute VB_UserMemId = 0
    ///     Foo = 42
    /// End Function
    ///
    /// Module:
    /// 
    /// Public Sub DoSomething(ByVal arg As Class1)
    ///     Dim bar As Variant
    ///     bar = arg.Foo(23)
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