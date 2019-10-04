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
    /// Identifies members of class modules that are used as interfaces, but that have a concrete implementation.
    /// </summary>
    /// <why>
    /// Interfaces provide an abstract, unified programmatic access to different objects; concrete implementations of their members should be in a separate module that 'Implements' the interface.
    /// </why>
    /// <example hasresult="false">
    /// <![CDATA[
    /// Option Explicit
    /// '@Interface
    ///
    /// Public Sub DoSomething()
    /// ' empty interface stub
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasresult="true">
    /// <![CDATA[
    /// Option Explicit
    /// '@Interface
    ///
    /// Public Sub DoSomething()
    ///     MsgBox "Hello from interface!"
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
                                                    result.QualifiedModuleName.ToString(),
                                                    Resources.RubberduckUI.ResourceManager
                                                        .GetString("DeclarationType_" + result.DeclarationType)
                                                        .Capitalize(),
                                                    result.IdentifierName),
                                        result));
        }
    }
}
