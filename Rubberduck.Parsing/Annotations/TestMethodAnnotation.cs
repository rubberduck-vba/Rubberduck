using System;
using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// Marks a method that the test engine will execute as a unit test.
    /// </summary>
    public sealed class TestMethodAnnotation : AnnotationBase
    {
        public TestMethodAnnotation(
            QualifiedSelection qualifiedSelection,
            IEnumerable<string> parameters)
            : base(AnnotationType.TestMethod, qualifiedSelection)
        {
            var firstParameter = parameters.FirstOrDefault();
            if ((firstParameter?.StartsWith("\"") ?? false) && firstParameter.EndsWith("\""))
            {
                // Strip surrounding double quotes
                firstParameter = firstParameter.Substring(1, firstParameter.Length - 2);
            }

            Category = string.IsNullOrWhiteSpace(firstParameter) ? string.Empty : firstParameter;
        }

        public string Category { get; }
    }
}
