using System.Collections.Generic;
using System.Linq;
using System.Windows.Data;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Refactorings.AnnotateDeclaration;
using Rubberduck.UI.Converters;

namespace Rubberduck.UI.Refactorings.AnnotateDeclaration
{
    internal class AnnotationArgumentViewModelFactory : IAnnotationArgumentViewModelFactory
    {
        private readonly IReadOnlyList<string> _inspectionNames;
        private readonly InspectionToLocalizedNameConverter _inspectionNameConverter;

        public AnnotationArgumentViewModelFactory(IEnumerable<IInspection> inspections, InspectionToLocalizedNameConverter inspectionNameConverter)
        {
            _inspectionNames = inspections
                .Select(inspection => inspection.AnnotationName)
                .ToList();
            _inspectionNameConverter = inspectionNameConverter;
        }
        
        public IAnnotationArgumentViewModel Create(AnnotationArgumentType argumentType, string argument = null)
        {
            var model = new TypedAnnotationArgument(argumentType, argument ?? string.Empty);
            return new AnnotationArgumentViewModel(model, _inspectionNames, _inspectionNameConverter);
        }
    }
}