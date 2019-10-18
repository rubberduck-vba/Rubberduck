using System.Linq;
using System.Threading;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class UntypedFunctionUsageInspectionTests
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

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected)
                .AddComponent("MyClass", ComponentType.ClassModule, inputCode)
                .AddReference("VBA", MockVbeBuilder.LibraryPathVBA, 4, 2, true)
                .Build();
            var vbe = builder.AddProject(project).Build();

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

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected)
                .AddComponent("MyClass", ComponentType.ClassModule, inputCode)
                .AddReference("VBA", MockVbeBuilder.LibraryPathVBA, 4, 2, true)
                .Build();
            var vbe = builder.AddProject(project).Build();

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

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected)
                .AddComponent("MyClass", ComponentType.ClassModule, inputCode)
                .AddReference("VBA", MockVbeBuilder.LibraryPathVBA, 4, 2, true)
                .Build();
            var vbe = builder.AddProject(project).Build();

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
            const string inspectionName = "UntypedFunctionUsageInspection";
            var inspection = new UntypedFunctionUsageInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}
