using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Inspections.Inspections.Extensions;
using Rubberduck.Common;
using System;
using Rubberduck.Parsing.Annotations;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Identifies class modules that define an interface with one or more members containing a concrete implementation.
    /// </summary>
    /// <why>
    /// Interfaces provide an abstract, unified programmatic access to different objects; concrete implementations of their members should be in a separate module that 'Implements' the interface.
    /// </why>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Option Explicit
    /// '@Interface
    ///
    /// Public Sub DoSomething()
    /// ' empty interface stub
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResults="true">
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
            var annotatedAsInterface = State.DeclarationFinder.Classes
                .Where(cls => cls.Annotations.Any(an => an.Annotation is InterfaceAnnotation)).Cast<ClassModuleDeclaration>();

            var implementedAndOrAnnotatedInterfaceModules = State.DeclarationFinder.FindAllUserInterfaces()
                .Union(annotatedAsInterface);

            return implementedAndOrAnnotatedInterfaceModules
                .SelectMany(interfaceModule => interfaceModule.Members
                    .Where(member => ((ModuleBodyElementDeclaration)member).Block.ContainsExecutableStatements(true)))
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
