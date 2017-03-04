using System.Linq;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Inspections;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.Application;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class UnassignedVariableUsageInspectionTests
    {
        [TestMethod]
        [TestCategory("Inspections")]
        public void UnassignedVariableUsage_ReturnsResult()
        {
            const string inputCode = 
@"Sub Foo()
    Dim b As Boolean
    Dim bb As Boolean
    bb = b
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected)
                .AddComponent("MyClass", ComponentType.ClassModule, inputCode)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new UnassignedVariableUsageInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        // this test will eventually be removed once we can fire the inspection on a specific reference
        [TestMethod]
        [TestCategory("Inspections")]
        public void UnassignedVariableUsage_ReturnsSingleResult_MultipleReferences()
        {
            const string inputCode =
@"Sub tester()
    Dim myarr() As Variant
    Dim i As Long

    ReDim myarr(1 To 10)

    For i = 1 To 10
        DoSomething myarr(i)
    Next

End Sub

Sub DoSomething(ByVal foo As Variant)
End Sub
";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected)
                .AddComponent("MyClass", ComponentType.ClassModule, inputCode)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new UnassignedVariableUsageInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnassignedVariableUsage_DoesNotReturnResult()
        {
            const string inputCode =
@"Sub Foo()
    Dim b As Boolean
    Dim bb As Boolean
    b = True
    bb = b
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected)
                .AddComponent("MyClass", ComponentType.ClassModule, inputCode)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new UnassignedVariableUsageInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.IsFalse(inspectionResults.Any());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnassignedVariableUsage_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
@"Sub Foo()
    '@Ignore UnassignedVariableUsage
    Dim b As Boolean
    Dim bb As Boolean

    bb = b
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected)
                .AddComponent("MyClass", ComponentType.ClassModule, inputCode)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new UnassignedVariableUsageInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.IsFalse(inspectionResults.Any());
        }

        [TestMethod]
        public void UnassignedVariableUsage_NoResultIfNoReferences()
        {
            const string inputCode =
@"Sub DoSomething()
    Dim foo
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected)
                .AddComponent("MyClass", ComponentType.ClassModule, inputCode)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new UnassignedVariableUsageInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.IsFalse(inspectionResults.Any());
        }

//        Ignored until we can reinstate the quick fix on a specific reference
//        [TestMethod]
//        [TestCategory("Inspections")]
//        public void UnassignedVariableUsage_QuickFixWorks()
//        {
//            const string inputCode =
//@"Sub Foo()
//    Dim b As Boolean
//    Dim bb As Boolean
//    bb = b
//End Sub";

//            const string expectedCode =
//@"Sub Foo()
//    Dim b As Boolean
//    Dim bb As Boolean
//    TODOTODO = TODO
//End Sub";

//            //Arrange
//            var builder = new MockVbeBuilder();
//            IVBComponent component;
//            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
//            var project = vbe.Object.VBProjects[0];
//            var module = project.VBComponents[0].CodeModule;
//            var mockHost = new Mock<IHostApplication>();
//            mockHost.SetupAllProperties();
//            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

//            parser.Parse(new CancellationTokenSource());
//            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

//            var inspection = new UnassignedVariableUsageInspection(parser.State);
//            var inspectionResults = inspection.GetInspectionResults();

//            inspectionResults.First().QuickFixes.First().Fix();
            
//            Assert.AreEqual(expectedCode, module.Content());
//        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnassignedVariableUsage_IgnoreQuickFixWorks()
        {
            const string inputCode =
@"Sub Foo()
    Dim b As Boolean
    Dim bb As Boolean
    bb = b
End Sub";

            const string expectedCode =
@"Sub Foo()
'@Ignore UnassignedVariableUsage
    Dim b As Boolean
    Dim bb As Boolean
    bb = b
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects[0];
            var module = project.VBComponents[0].CodeModule;
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new UnassignedVariableUsageInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            inspectionResults.First().QuickFixes.Single(s => s is IgnoreOnceQuickFix).Fix();

            Assert.AreEqual(expectedCode, module.Content());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionType()
        {
            var inspection = new UnassignedVariableUsageInspection(null);
            Assert.AreEqual(CodeInspectionType.CodeQualityIssues, inspection.InspectionType);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = "UnassignedVariableUsageInspection";
            var inspection = new UnassignedVariableUsageInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}
