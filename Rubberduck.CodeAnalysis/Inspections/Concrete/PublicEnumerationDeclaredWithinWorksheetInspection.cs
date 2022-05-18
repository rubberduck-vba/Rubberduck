using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;
using Rubberduck.VBEditor.SafeComWrappers;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Identifies public enumerations declared within worksheet modules.
    /// </summary>
    /// <why>
    /// Copying a worksheet which contains a public `Enum` declaration will duplicate the enum resulting in a state which prevents compilation.
    /// </why>
    /// <example hasResult="true">
    /// <module name="DocumentModule" type="Document Module">
    /// <![CDATA[
    /// Public Enum Foo()
    ///     ' enumeration members
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="DocumentModule" type="Document Module">
    /// <![CDATA[
    /// Private Enum Foo()
    ///     ' enumeration members
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class PublicEnumerationDeclaredWithinWorksheetInspection : DeclarationInspectionBase
    {
        private readonly string[] _worksheetSuperTypeNames = new string[] { "Worksheet", "_Worksheet" };

        public PublicEnumerationDeclaredWithinWorksheetInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider, DeclarationType.Enumeration)
        {}

        protected override bool IsResultDeclaration(Declaration enumeration, DeclarationFinder finder)
        {
            if (enumeration.Accessibility != Accessibility.Private
               && enumeration.QualifiedModuleName.ComponentType == ComponentType.Document)
            {
                if (enumeration.ParentDeclaration is ClassModuleDeclaration classModuleDeclaration)
                {
                    return RetrieveSuperTypeNames(classModuleDeclaration).Intersect(_worksheetSuperTypeNames).Any();
                }
            }

            return false;
        }

        protected override string ResultDescription(Declaration declaration)
        {
            return string.Format(InspectionResults.PublicEnumerationDeclaredWithinWorksheetInspection,
                declaration.IdentifierName,
                declaration.ParentScopeDeclaration.IdentifierName);
        }

        /// <summary>
        /// Supports property injection for testing. 
        /// </summary>
        /// <remarks>
        /// MockParser does not populate SuperTypes/SuperTypeNames.  RetrieveSuperTypeNames Func allows injection
        /// of ClassModuleDecularation.SuperTypeNames property results.
        /// </remarks>
        public Func<ClassModuleDeclaration, IEnumerable<string>> RetrieveSuperTypeNames { set; private get; } = GetSuperTypeNames;

        private static IEnumerable<string> GetSuperTypeNames(ClassModuleDeclaration classModule)
        {
            return classModule.SupertypeNames;
        }
    }
}
