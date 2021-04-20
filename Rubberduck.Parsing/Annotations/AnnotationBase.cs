using Rubberduck.VBEditor.SafeComWrappers;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Parsing.Annotations
{
    public abstract class AnnotationBase : IAnnotation
    {
        public bool AllowMultiple { get; }
        public int RequiredArguments { get; }
        public int? AllowedArguments { get; }
        public IReadOnlyList<AnnotationArgumentType> AllowedArgumentTypes { get; }
        public string Name { get; }
        public AnnotationTarget Target { get; }
        public virtual IReadOnlyList<ComponentType> IncompatibleComponentTypes { get; } = new ComponentType[] { };
        public virtual ComponentType? RequiredComponentType { get; } = null;

        protected AnnotationBase(string name, AnnotationTarget target, 
            int requiredArguments = 0, 
            int? allowedArguments = 0, 
            IReadOnlyList<AnnotationArgumentType> allowedArgumentTypes = null, 
            bool allowMultiple = false)
        {
            Name = name;
            Target = target;
            AllowMultiple = allowMultiple;
            RequiredArguments = requiredArguments;
            AllowedArguments = allowedArguments;
            AllowedArgumentTypes = allowedArgumentTypes ?? new List<AnnotationArgumentType>();
        }

        public virtual IReadOnlyList<string> ProcessAnnotationArguments(IEnumerable<string> arguments) => arguments.ToList();

        public override bool Equals(object obj) => obj is AnnotationBase annotation && Equals(annotation);

        public bool Equals(AnnotationBase other) => other != null && Name.Equals(other.Name);
        
        public override int GetHashCode() => Name.GetHashCode();
    }
}
