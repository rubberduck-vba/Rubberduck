using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Inspections.Inspections.Extensions;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Warns about properties that don't expose a 'Property Get' accessor.
    /// </summary>
    /// <why>
    /// Write-only properties are suspicious: if the client code is able to set a property, it should be allowed to read that property as well. 
    /// Class design guidelines and best practices generally recommend against write-only properties.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Private internalFoo As Long
    ///
    /// Public Property Let Foo(ByVal value As Long)
    ///     internalFoo = value
    /// End Property
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Private internalFoo As Long
    ///
    /// Public Property Let Foo(ByVal value As Long)
    ///     internalFoo = value
    /// End Property
    ///
    /// Public Property Get Foo() As Long
    ///     Foo = internalFoo
    /// End Property
    /// ]]>
    /// </example>
    public sealed class WriteOnlyPropertyInspection : InspectionBase
    {
        public WriteOnlyPropertyInspection(RubberduckParserState state)
            : base(state) { }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var setters = State.DeclarationFinder.UserDeclarations(DeclarationType.Property | DeclarationType.Procedure)
                .Where(item => 
                       (item.Accessibility == Accessibility.Implicit || 
                        item.Accessibility == Accessibility.Public || 
                        item.Accessibility == Accessibility.Global)
                    && State.DeclarationFinder.MatchName(item.IdentifierName).All(accessor => accessor.DeclarationType != DeclarationType.PropertyGet))
                .GroupBy(item => new {item.QualifiedName, item.DeclarationType})
                .Select(grouping => grouping.First()); // don't get both Let and Set accessors

            return setters.Select(setter =>
                new DeclarationInspectionResult(this,
                                                string.Format(InspectionResults.WriteOnlyPropertyInspection, setter.IdentifierName),
                                                setter));
        }
    }
}
