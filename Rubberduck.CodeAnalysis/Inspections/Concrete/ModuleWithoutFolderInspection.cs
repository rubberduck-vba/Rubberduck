using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.Annotations;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class ModuleWithoutFolderInspection : InspectionBase
    {
        public ModuleWithoutFolderInspection(RubberduckParserState state)
            : base(state)
        {
        }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var modulesWithoutFolderAnnotation = State.DeclarationFinder.UserDeclarations(Parsing.Symbols.DeclarationType.Module)
                .Where(w => w.Annotations.All(a => a.AnnotationType != AnnotationType.Folder))
                .ToList();

            return modulesWithoutFolderAnnotation.Select(declaration =>
                new DeclarationInspectionResult(this, string.Format(InspectionResults.ModuleWithoutFolderInspection, declaration.IdentifierName), declaration));
        }
    }
}
