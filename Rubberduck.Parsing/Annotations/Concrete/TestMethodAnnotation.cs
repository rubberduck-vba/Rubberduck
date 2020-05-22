using System.Collections.Generic;
using System.Linq;
using Rubberduck.Common;

namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// Marks a method that the test engine will execute as a unit test.
    /// </summary>
    public sealed class TestMethodAnnotation : AnnotationBase, ITestAnnotation
    {
        public TestMethodAnnotation()
            : base("TestMethod", AnnotationTarget.Member)
        {}

        public IReadOnlyList<string> ProcessAnnotationArguments(IEnumerable<string> arguments)
        {
            var firstParameter = arguments.FirstOrDefault();
            var result = new List<string>();
            if (!string.IsNullOrWhiteSpace(firstParameter))
            {
                result.Add(firstParameter);
            }
            return result;
        }
    }
}
