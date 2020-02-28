using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Parsing.Annotations
{
    public abstract class AnnotationBase : IAnnotation
    {
        public bool AllowMultiple { get; }
        public int RequiredArguments { get; }
        public int? AllowedArguments { get; }
        public string Name { get; }
        public AnnotationTarget Target { get; }

        protected AnnotationBase(string name, AnnotationTarget target, int requiredArguments = 0, int? allowedArguments = 0, bool allowMultiple = false)
        {
            Name = name;
            Target = target;
            AllowMultiple = allowMultiple;
            RequiredArguments = requiredArguments;
            AllowedArguments = allowedArguments;
        }

        public virtual IReadOnlyList<string> ProcessAnnotationArguments(IEnumerable<string> arguments)
        {
            return arguments.ToList();
        }
    }
}
