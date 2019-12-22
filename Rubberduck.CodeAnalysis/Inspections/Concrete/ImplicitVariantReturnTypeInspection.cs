using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Inspections.Inspections.Extensions;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Warns about 'Function' and 'Property Get' procedures that don't have an explicit return type.
    /// </summary>
    /// <why>
    /// All functions return something, whether a type is specified or not. The implicit default is 'Variant'.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Public Function GetFoo()
    ///     GetFoo = 42
    /// End Function
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Public Function GetFoo() As Long
    ///     GetFoo = 42
    /// End Function
    /// ]]>
    /// </example>
    public sealed class ImplicitVariantReturnTypeInspection : InspectionBase
    {
        public ImplicitVariantReturnTypeInspection(RubberduckParserState state)
            : base(state) { }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var issues = from item in State.DeclarationFinder.UserDeclarations(DeclarationType.Function)
                         where !item.IsTypeSpecified 
                         let issue = new {Declaration = item, QualifiedContext = new QualifiedContext<ParserRuleContext>(item.QualifiedName, item.Context)}
                         select new DeclarationInspectionResult(this,
                                                     string.Format(InspectionResults.ImplicitVariantReturnTypeInspection, item.IdentifierName),
                                                     item);
            return issues;
        }
    }
}
