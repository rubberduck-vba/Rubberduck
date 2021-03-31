using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Parsing.Annotations.Concrete
{
    /// <summary>
    /// @EntryPoint annotation, marks a procedure as an entry point that is not intended to be referenced directly from the code.
    /// </summary>
    /// <parameter name="CallerName" type="Text" required="False">
    /// If provided, the first argument is interpreted as referring to an external caller, for example the name of a Shape in the host document.
    /// </parameter>
    /// <remarks>
    /// Members with this annotation are ignored by the ProcedureNotUsed inspection. The CallerName argument is currently not being validated, but may be in the future.
    /// </remarks>
    /// <example>
    /// <module name="Module1" type="Standard Module">
    /// </module>
    /// <![CDATA[
    /// Option Explicit
    ///
    /// '@EntryPoint "Rounded Rectangle 1"
    /// Public Sub DoSomething()
    /// End Sub
    /// 
    /// '@EntryPoint "Rounded Rectangle 2"
    /// Public Sub DoSomethingElse()
    /// End Sub
    /// ]]>
    /// </example>
    public sealed class EntryPointAnnotation : AnnotationBase
    {
        public string CallerName { get; private set; }

        public EntryPointAnnotation()
            : base("EntryPoint", AnnotationTarget.Member, allowedArguments: 1, allowedArgumentTypes: new[] { AnnotationArgumentType.Text })
        { }

        public override IReadOnlyList<string> ProcessAnnotationArguments(IEnumerable<string> arguments)
        {
            var args = arguments.ToList();

            CallerName = args.Any()
                ? args[0]
                : string.Empty;

            return base.ProcessAnnotationArguments(args);
        }
    }
}