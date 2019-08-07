using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Inspections.Inspections.Extensions;
using Rubberduck.Common;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Identifies implemented members of class modules that are used as interfaces.
    /// </summary>
    /// <why>
    /// Interfaces provide a unified programmatic access to different objects, and therefore are rearly instantiated as concrete objects.
    /// </why>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Sub Foo()
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Sub Foo()
    ///     MsgBox "?"
    /// End Sub
    /// ]]>
    /// </example>
    internal class ImplementedInterfaceMemberInspection : InspectionBase
    {
        public ImplementedInterfaceMemberInspection(Parsing.VBA.RubberduckParserState state)
            : base(state) { }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            return State.DeclarationFinder.FindAllUserInterfaces()
                .SelectMany(interfaceModule => interfaceModule.Members
                    .Where(member => ((ModuleBodyElementDeclaration)member).Block.ContainsExecutableStatements(true)
                                     && !member.IsIgnoringInspectionResultFor(AnnotationName)))
                .Select(result => new DeclarationInspectionResult(this,
                                        string.Format(InspectionResults.ImplementedInterfaceMemberInspection,
                                                    Resources.RubberduckUI.ResourceManager
                                                        .GetString("DeclarationType_" + result.DeclarationType)
                                                        .Capitalize(),
                                                    result.IdentifierName),
                                        result));
        }
    }
}