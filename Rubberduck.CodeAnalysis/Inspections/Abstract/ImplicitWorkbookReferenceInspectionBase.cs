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

        private IReadOnlyList<ModuleDeclaration> _relevantClasses;
        private IReadOnlyList<PropertyGetDeclaration> _relevantProperties;

        protected Declaration Excel { get; private set; }

        protected override IEnumerable<Declaration> ObjectionableDeclarations(DeclarationFinder finder)
        {
            if (Excel == null)
            {
                if (!finder.TryFindProjectDeclaration("Excel", out var excel))
                {
                    return Enumerable.Empty<Declaration>();
                }
                Excel = excel;
            }

            if (_relevantClasses == null)
            {
                _relevantClasses = InterestingClasses
                    .Select(className => finder.FindClassModule(className, Excel, true))
                    .OfType<ModuleDeclaration>()
                    .ToList();
            }

            if (_relevantProperties == null)
            {
                _relevantProperties = _relevantClasses
                    .SelectMany(classDeclaration => classDeclaration.Members)
                    .OfType<PropertyGetDeclaration>()
                    .Where(member => InterestingMembers.Contains(member.IdentifierName))
                    .ToList();
            }

            return _relevantProperties;
        }
    }
}