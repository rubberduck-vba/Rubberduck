using Rubberduck.Parsing.Annotations;

namespace Rubberduck.UI.Refactorings.AnnotateDeclaration
{
    internal interface IAnnotationArgumentViewModelFactory
    {
        IAnnotationArgumentViewModel Create(AnnotationArgumentType argumentType, string argument = null);
    }
}