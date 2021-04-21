using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;

namespace Rubberduck.CodeAnalysis.Inspections.Abstract
{
    internal abstract class ImplicitSheetReferenceInspectionBase : IdentifierReferenceInspectionFromDeclarationsBase
    {
        public ImplicitSheetReferenceInspectionBase(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        { }

        protected override IEnumerable<Declaration> ObjectionableDeclarations(DeclarationFinder finder)
        {
            var excel = finder.Projects.SingleOrDefault(item => !item.IsUserDefined && item.IdentifierName == "Excel");
            if (excel == null)
            {
                return Enumerable.Empty<Declaration>();
            }

            var globalModules = GlobalObjectClassNames
                .Select(className => finder.FindClassModule(className, excel, true))
                .OfType<ModuleDeclaration>();

            return globalModules
                .SelectMany(moduleClass => moduleClass.Members)
                .Where(declaration => TargetMemberNames.Contains(declaration.IdentifierName)
                                      && declaration.DeclarationType.HasFlag(DeclarationType.Member)
                                      && declaration.AsTypeName == "Range");
        }

        protected virtual string[] GlobalObjectClassNames => new[]
        {
            "Global", "_Global", 
            "Worksheet", "_Worksheet"
        };

        protected virtual string[] TargetMemberNames => new[]
        {
            "Cells", "Range", "Columns", "Rows"
        };
    }
}