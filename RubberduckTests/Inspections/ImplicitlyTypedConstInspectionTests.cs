using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class ImplicitlyTypedConstInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            var inspectionUnderTest = new ImplicitlyTypedConstInspection(null);

            var expected = nameof(ImplicitlyTypedConstInspection);
            var actual = inspectionUnderTest.Name;
            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("Inspections")]
        public void ImplicitlyTypedConst_ReturnsResult()
        {
            const string inputCode =
@"Public Sub Foo()
    Const bar = 0
End Sub";

            const int expected = 1;

            var results = InspectionResultsForModules(("FooClass", inputCode, ComponentType.ClassModule));
            var actual = results.Count();
            Assert.AreEqual(expected , actual);
        }

        [Test]
        [Category("Inspections")]
        public void ImplicitlyTypedConst_DoesNotReturnResult()
        {
            const string inputCode =
@"Public Sub Foo()
    Const bar As Long = 0
End Sub";

            const int expected = 0;

            var results = InspectionResultsForModules(("FooClass", inputCode, ComponentType.ClassModule));
            var actual = results.Count();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("Inspections")]
        [TestCase("!")]
        [TestCase("@")]
        [TestCase("#")]
        [TestCase("$")]
        [TestCase("%")]
        [TestCase("^")]
        [TestCase("&")]
        public void ImplicitlyTypedConst_HasTypeHint_DoesNotReturnResult(string typeHint)
        {
            string inputCode =
$@"Public Sub Foo()
    Const bar{typeHint} = 0
End Sub";

            const int expected = 0;

            var results = InspectionResultsForModules(("FooClass", inputCode, ComponentType.ClassModule));
            var actual = results.Count();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("Inspections")]
        [TestCase("!")]
        [TestCase("@")]
        [TestCase("#")]
        //[TestCase("$")] Errors out
        [TestCase("%")]
        [TestCase("^")]
        [TestCase("&")]
        public void ImplicitlyTypedConst_AssignedValueHasTypeHint_ReturnsResult(string typeHint)
        {
            string inputCode =
$@"Public Sub Foo()
    Const bar = 0{typeHint}
End Sub";

            const int expected = 1;

            var results = InspectionResultsForModules(("FooClass", inputCode, ComponentType.ClassModule));
            var actual = results.Count();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("Inspections")]
        public void ImplicitlyTypedConst_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
@"Sub Foo()
    '@Ignore ImplicitlyTypedConst
    Const bar = 0
End Sub";

            const int expected = 0;

            var results = InspectionResultsForModules(("FooClass", inputCode, ComponentType.ClassModule));
            var actual = results.Count();

            Assert.AreEqual(expected, actual);
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new ImplicitlyTypedConstInspection(state);
        }
    }
}
