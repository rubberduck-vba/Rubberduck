using System.Collections.Generic;
using System.Linq;
using Rubberduck.Common;

namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// Used to mark members as obsolete, so that Rubberduck can warn users whenever they try to use an obsolete member.
    /// </summary>
    public sealed class ObsoleteAnnotation : AnnotationBase
    {
        public string ReplacementDocumentation { get; private set; }

        public ObsoleteAnnotation()
            : base("Obsolete", AnnotationTarget.Member | AnnotationTarget.Variable, allowedArguments: 1)
        {}

        public override IReadOnlyList<string> ProcessAnnotationArguments(IEnumerable<string> arguments)
        {
            var args = arguments.ToList();

            ReplacementDocumentation = args.Any()
                ? args[0]
                : string.Empty;

            return base.ProcessAnnotationArguments(args);
        }
    }
}