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
    /// Copying a worksheet which contains a public Enum declaration will also create a copy of the Enum declaration.  
    /// The copied Enum declaration will result in an 'Ambiguous name detected' compiler error.  
    /// Declaring Enumerations in Standard or Class modules avoids unintentional duplication of an Enum declaration.
    /// </why>
    /// <example hasResult="true">
    /// <module name="WorksheetModule" type="Document Module">
    /// <![CDATA[
    /// Public Enum ExampleEnum
    ///     FirstEnum = 0
    ///     SecondEnum
    /// End Enum
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="WorksheetModule" type="Document Module">
    /// <![CDATA[
    /// Private Enum ExampleEnum
    ///     FirstEnum = 0
    ///     SecondEnum
    /// End Enum
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class PublicEnumerationDeclaredInWorksheetInspection : DeclarationInspectionBase
    {
        private readonly string[] _worksheetSuperTypeNames = new string[] { "Worksheet", "_Worksheet" };

        public PublicEnumerationDeclaredInWorksheetInspection(IDeclarationFinderProvider declarationFinderProvider)
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
            return string.Format(InspectionResults.PublicEnumerationDeclaredInWorksheetInspection,
                declaration.IdentifierName);
        }

        /// <summary>
        /// Supports property injection for testing. 
        /// </summary>
        /// <remarks>
        /// MockParser does not populate SuperTypes/SuperTypeNames.  RetrieveSuperTypeNames Func allows injection
        /// of ClassModuleDeclaration.SuperTypeNames property results.
        /// </remarks>
        public Func<ClassModuleDeclaration, IEnumerable<string>> RetrieveSuperTypeNames { set; private get; } = GetSuperTypeNames;

        private static IEnumerable<string> GetSuperTypeNames(ClassModuleDeclaration classModule)
        {
            return classModule.SupertypeNames;
        }
    }
}
