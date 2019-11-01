using System.Linq;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.Parsing.Inspections.Abstract;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class UntypedFunctionUsageInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void UntypedFunctionUsage_ReturnsResult()
        {
            const string inputCode =
@"Sub Foo()
    Dim str As String
    str = Left(""test"", 1)
End Sub";

            Assert.AreEqual(1, InspectionResultsForModules(("MyClass", inputCode, ComponentType.ClassModule), "VBA").Count());
        }

        [Test]
        [Category("Inspections")]
        public void UntypedFunctionUsage_DoesNotReturnResult()
        {
            const string inputCode =
@"Sub Foo()
    Dim str As String
    str = Left$(""test"", 1)
End Sub";

            Assert.AreEqual(0, InspectionResultsForModules(("MyClass", inputCode, ComponentType.ClassModule), "VBA").Count());
        }

        [Test]
        [Category("Inspections")]
        public void UntypedFunctionUsage_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
@"Sub Foo()
    Dim str As String

    '@Ignore UntypedFunctionUsage
    str = Left(""test"", 1)
End Sub";

            Assert.AreEqual(0, InspectionResultsForModules(("MyClass", inputCode, ComponentType.ClassModule), "VBA").Count());
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            var inspection = new UntypedFunctionUsageInspection(null);

            Assert.AreEqual(nameof(UntypedFunctionUsageInspection), inspection.Name);
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new UntypedFunctionUsageInspection(state);
        }
    }
}
