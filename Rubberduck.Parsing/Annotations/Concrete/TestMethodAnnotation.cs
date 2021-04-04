using System.Collections.Generic;
using System.Linq;
using Rubberduck.Common;
using Rubberduck.Parsing.Annotations;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.Parsing.Annotations.Concrete
{
    /// <summary>
    /// @TestMethod annotation, identifies a procedure that constitutes a unit test.
    /// </summary>
    /// <parameter name="TestCategory" type="Text">
    /// If arguments are supplied, the current implementation makes the first provided argument be the test category.
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
            : base("TestMethod", AnnotationTarget.Member, allowedArguments: 1, allowedArgumentTypes: new []{AnnotationArgumentType.Text})
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

        public override ComponentType? RequiredComponentType => ComponentType.StandardModule;
    }
}
