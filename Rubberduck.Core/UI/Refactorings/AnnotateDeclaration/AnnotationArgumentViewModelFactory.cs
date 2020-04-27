using System.Collections.Generic;
using System.Linq;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Refactorings.AnnotateDeclaration;

namespace Rubberduck.UI.Refactorings.AnnotateDeclaration
{
    internal class AnnotationArgumentViewModelFactory : IAnnotationArgumentViewModelFactory
    {
        private readonly IReadOnlyList<string> InspectionNames;

        public AnnotationArgumentViewModelFactory(IEnumerable<IInspection> inspections)
        {
            InspectionNames = inspections
                .Select(inspection => inspection.AnnotationName)
                .ToList();
        }
        
        public IAnnotationArgumentViewModel Create(AnnotationArgumentType argumentType, string argument = null)
        {
            var model = new TypedAnnotationArgument(argumentType, argument ?? string.Empty);
            return new AnnotationArgumentViewModel(model, InspectionNames);
        }
    }
}