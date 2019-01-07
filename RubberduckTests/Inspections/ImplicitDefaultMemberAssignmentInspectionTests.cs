using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class ImplicitDefaultMemberAssignmentInspectionTests
    {
        [Test]
        [Category("Inspections")]
        public void ImplicitDefaultMemberAssignment_ReturnsResult()
        {
            const string inputCode =
@"Public Sub Foo(bar As Range)
    With bar
        .Cells(1, 1) = 42
    End With
End Sub
";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", "TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, inputCode)
                .AddReference("Excel", MockVbeBuilder.LibraryPathMsExcel, 1, 8, true)
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

                var inspection = new ImplicitDefaultMemberAssignmentInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void ImplicitDefaultMemberAssignment_IgnoredDoesNotReturnResult()
        {
            const string inputCode =
@"Public Sub Foo(bar As Range)
    With bar
        '@Ignore ImplicitDefaultMemberAssignment
        .Cells(1, 1) = 42
    End With
End Sub
";
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", "TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, inputCode)
                .AddReference("Excel", MockVbeBuilder.LibraryPathMsExcel, 1, 8, true)
                .Build();
            var vbe = builder.AddProject(project).Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ImplicitDefaultMemberAssignmentInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void ImplicitDefaultMemberAssignment_ExplicitCallDoesNotReturnResult()
        {
            const string inputCode =
@"Public Sub Foo(bar As Range)
    With bar
        .Cells(1, 1).Value = 42
    End With
End Sub
";
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", "TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, inputCode)
                .AddReference("Excel", MockVbeBuilder.LibraryPathMsExcel, 1, 8, true)
                .Build();
            var vbe = builder.AddProject(project).Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ImplicitDefaultMemberAssignmentInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = "ImplicitDefaultMemberAssignmentInspection";
            var inspection = new ImplicitDefaultMemberAssignmentInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}
