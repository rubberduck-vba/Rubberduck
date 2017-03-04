using System.Linq;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Inspections;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.Application;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class HungarianNotationInspectionTests
    {
        [TestMethod]
        [TestCategory("Inspections")]
        public void HungarianNotation_ReturnsResult_VariableWithThreeLetterPrefix()
        {
            const string inputCode =
@"Sub Hungarian()
    Dim strFoo As String
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

            var inspection = new HungarianNotationInspection(null, parser.State, UseMeaningfulNameInspectionTests.GetInspectionSettings().Object);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void HungarianNotation_ReturnsResult_VariableWithOneLetterPrefix()
        {
            const string inputCode =
@"Sub Hungarian()
    Dim oFoo As Object
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

            var inspection = new HungarianNotationInspection(null, parser.State, UseMeaningfulNameInspectionTests.GetInspectionSettings().Object);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void HungarianNotation_ReturnsResult_ForClass()
        {
            const string inputCode =
@"Sub Test()
    Debug.Print ""Ez egy objektum""
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected)
                .AddComponent("clsMagyar", ComponentType.ClassModule, inputCode)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new HungarianNotationInspection(null, parser.State, UseMeaningfulNameInspectionTests.GetInspectionSettings().Object);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void HungarianNotation_DoesNotReturnsResult_AllLowerCase()
        {
            const string inputCode =
@"Sub NoHungarianHere()
    Dim strong As Variant
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

            var inspection = new HungarianNotationInspection(null, parser.State, UseMeaningfulNameInspectionTests.GetInspectionSettings().Object);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.IsFalse(inspectionResults.Any());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void HungarianNotation_DoesNotReturnsResult_UpperCaseFirstLetter()
        {
            const string inputCode =
@"Option Explicit";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected)
                .AddComponent("Stripper", ComponentType.ClassModule, inputCode)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new HungarianNotationInspection(null, parser.State, UseMeaningfulNameInspectionTests.GetInspectionSettings().Object);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.IsFalse(inspectionResults.Any());
        }
        
        [TestMethod]
        [TestCategory("Inspections")]
        public void HungarianNotation_DoesNotReturnsResult_ThreeLetterVariable()
        {
            const string inputCode =
@"Sub InExcelSomewhere()
    Dim col As Long
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

            var inspection = new HungarianNotationInspection(null, parser.State, UseMeaningfulNameInspectionTests.GetInspectionSettings().Object);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.IsFalse(inspectionResults.Any());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void HungarianNotation_DoesNotReturnResult_WhenIgnored()
        {
            const string inputCode =
@"Sub MagyarRendbenVan()
    '@Ignore HungarianNotation
    Dim strFoo As Variant
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

            var inspection = new UseMeaningfulNameInspection(null, parser.State, UseMeaningfulNameInspectionTests.GetInspectionSettings().Object);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.IsFalse(inspectionResults.Any());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void HungarianNotation_DoesNotReturnResult_WhenWhitelisted()
        {
            const string inputCode =
@"Sub Feherlista()
    Dim oRange As Object
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

            var inspection = new UseMeaningfulNameInspection(null, parser.State, UseMeaningfulNameInspectionTests.GetInspectionSettings().Object);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.IsFalse(inspectionResults.Any());
        }
    }
}
