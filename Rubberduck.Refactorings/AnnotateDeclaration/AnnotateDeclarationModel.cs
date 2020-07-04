using System.Collections.Generic;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Refactorings.AnnotateDeclaration
{
    //This is introduced in favor of a value tuple in order to be able to bind to the components in XAML.
    public struct TypedAnnotationArgument
    {
        public AnnotationArgumentType ArgumentType { get; set; }
        public string Argument { get; set; }

        public TypedAnnotationArgument(AnnotationArgumentType type, string argument)
        {
            ArgumentType = type;
            Argument = argument;
        }
    }

    public class AnnotateDeclarationModel : IRefactoringModel
    {
        public Declaration Target { get; }
        public IAnnotation Annotation { get; set; }
        public IList<TypedAnnotationArgument> Arguments { get; set; }
        public bool AdjustAttribute { get; set; }

        public AnnotateDeclarationModel(
            Declaration target, 
            IAnnotation annotation = null,
            IList<TypedAnnotationArgument> arguments = null,
            bool adjustAttribute = false)
        {
            Target = target;
            Annotation = annotation;
            Arguments = arguments ?? new List<TypedAnnotationArgument>();
            AdjustAttribute = adjustAttribute;
        }
    }
}