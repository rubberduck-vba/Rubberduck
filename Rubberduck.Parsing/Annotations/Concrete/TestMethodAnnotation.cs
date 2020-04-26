using System.Collections.Generic;
using System.Linq;
using Rubberduck.Common;

namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// @TestMethod annotation, identifies procedures that contain unit tests.
    /// </summary>
    /// <parameter>
    /// This annotation takes an optional string argument identifying the test category.
    /// </parameter>
    /// <example>
    /// <module name="TestModule1" type="Standard Module">
    /// <![CDATA[
    /// Option Explicit
    /// '@TestModule
    /// 
    /// Private Assert As Rubberduck.AssertClass
    /// '...
    /// 
    /// '@TestMethod("Category")
    /// Private Sub GivenThing_ThenResult()
    ///     'use Assert calls to specify conditions that make the test fail.
    ///     Assert.IsTrue False
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
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
