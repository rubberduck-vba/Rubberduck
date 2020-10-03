using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// @Obsolete annotation, marks a procedure as "obsolete". Rubberduck inspections can then warn about code that references them.
    /// </summary>
    /// <parameter name="Message" type="Text" required="False">
    /// If provided, the first argument becomes additional metadata that Rubberduck can use in inspection results, and remains a valid and useful in-place comment.
    /// </parameter>
    /// <example>
    /// <module name="Class1" type="Class Module">
    /// <![CDATA[
    /// Option Explicit
    ///
    /// '@Obsolete("Use DoSomethingElse instead.")
    /// Public Sub DoSomething()
    /// End Sub
    /// 
    /// Public Sub DoSomethingElse()
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    public sealed class ObsoleteAnnotation : AnnotationBase
    {
        public string ReplacementDocumentation { get; private set; }

        public ObsoleteAnnotation()
            : base("Obsolete", AnnotationTarget.Member | AnnotationTarget.Variable, allowedArguments: 1, allowedArgumentTypes: new [] {AnnotationArgumentType.Text})
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