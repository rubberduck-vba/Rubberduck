using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;

namespace Rubberduck.CodeAnalysis.Inspections.Abstract
{
    internal abstract class ImplicitWorkbookReferenceInspectionBase : IdentifierReferenceInspectionFromDeclarationsBase
    {
        internal ImplicitWorkbookReferenceInspectionBase(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        { }

        protected virtual string[] InterestingMembers => new[]
        {
            "Worksheets", "Sheets", "Names"
        };

        protected virtual string[] InterestingClasses => new[]
        {
            "_Global", "_Application", "Global", "Application", "_Workbook", "Workbook"
        };

        protected Declaration Excel(DeclarationFinder finder)
        {
            return finder.BuiltInDeclarations(DeclarationType.Project)
                .FirstOrDefault(project => project.IdentifierName.Equals("Excel", StringComparison.InvariantCultureIgnoreCase));
        } 

        protected override IEnumerable<Declaration> ObjectionableDeclarations(DeclarationFinder finder)
        {
            var excel = Excel(finder);
            var relevantClasses = InterestingClasses
                .Select(className => finder.FindClassModule(className, excel, true))
                .OfType<ModuleDeclaration>()
                .ToList();

            var relevantProperties = relevantClasses
                .SelectMany(classDeclaration => classDeclaration.Members)
                .OfType<PropertyGetDeclaration>()
                .Where(member => InterestingMembers.Contains(member.IdentifierName))
                .ToList();

            return relevantProperties;
        }
    }
}