using System.Linq;
using System.Threading;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;
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
            var vbe = MockVbeBuilder.BuildFromModules(("MyClass", inputCode, ComponentType.ClassModule),"VBA");

            var parser = MockParser.Create(vbe.Object);
            using (var state = parser.State)
            {
                parser.Parse(new CancellationTokenSource());
                if (state.Status >= ParserState.Error)
                {
                    Assert.Inconclusive("Parser Error");
                }

                var inspection = new UntypedFunctionUsageInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(1, inspectionResults.Count());
            }
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

            var vbe = MockVbeBuilder.BuildFromModules(("MyClass", inputCode, ComponentType.ClassModule));

            var parser = MockParser.Create(vbe.Object);
            using (var state = parser.State)
            {
                parser.Parse(new CancellationTokenSource());
                if (state.Status >= ParserState.Error)
                {
                    Assert.Inconclusive("Parser Error");
                }

                var inspection = new UntypedFunctionUsageInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.IsFalse(inspectionResults.Any());
            }
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

            var vbe = MockVbeBuilder.BuildFromModules(("MyClass", inputCode, ComponentType.ClassModule), "VBA");

            var parser = MockParser.Create(vbe.Object);
            using (var state = parser.State)
            {
                parser.Parse(new CancellationTokenSource());
                if (state.Status >= ParserState.Error)
                {
                    Assert.Inconclusive("Parser Error");
                }

                var inspection = new UntypedFunctionUsageInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.IsFalse(inspectionResults.Any());
            }
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
