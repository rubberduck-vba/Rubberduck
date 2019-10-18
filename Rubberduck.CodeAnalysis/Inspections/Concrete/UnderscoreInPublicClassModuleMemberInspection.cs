using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.VBA;
using Rubberduck.Inspections.Inspections.Extensions;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Warns about public class members with an underscore in their names.
    /// </summary>
    /// <why>
    /// The public interface of any class module can be implemented by any other class module; if the public interface 
    /// contains names with underscores, other classes cannot implement it - the code will not compile. Avoid underscores; prefer PascalCase names.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// '@Interface
    /// 
    /// Public Sub Do_Something() ' underscore in name makes the interface non-implementable.
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// '@Interface
    /// 
    /// Public Sub DoSomething() ' PascalCase identifiers are never a problem.
    /// End Sub
    /// ]]>
    /// </example>
    public sealed class UnderscoreInPublicClassModuleMemberInspection : InspectionBase
    {
        public UnderscoreInPublicClassModuleMemberInspection(RubberduckParserState state)
            : base(state) { }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var interfaceMembers = State.DeclarationFinder.FindAllInterfaceImplementingMembers().ToList();
            var eventHandlers = State.DeclarationFinder.FindEventHandlers().ToList();

            var names = State.DeclarationFinder.UserDeclarations(Parsing.Symbols.DeclarationType.Member)
                .Where(w => w.ParentDeclaration.DeclarationType.HasFlag(Parsing.Symbols.DeclarationType.ClassModule))
                .Where(w => !interfaceMembers.Contains(w) && !eventHandlers.Contains(w))
                .Where(w => w.Accessibility == Parsing.Symbols.Accessibility.Public || w.Accessibility == Parsing.Symbols.Accessibility.Implicit)
                .Where(w => w.IdentifierName.Contains('_'))
                .ToList();

            return names.Select(issue =>
                new DeclarationInspectionResult(this, string.Format(InspectionResults.UnderscoreInPublicClassModuleMemberInspection, issue.IdentifierName), issue));
        }
    }
}
