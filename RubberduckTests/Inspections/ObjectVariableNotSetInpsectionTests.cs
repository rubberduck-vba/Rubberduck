using System.Linq;
using Microsoft.Vbe.Interop;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Inspections;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.VBEHost;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class ObjectVariableNotSetInpsectionTests
    {
        [TestMethod]
        public void ObjectVariableNotSet_GivenStringVariable_ReturnsNoResult()
        {
            const string inputCode = @"
Private Sub Workbook_Open()
    
    Dim target As String
    target = Range(""A1"")
    
    target.Value = ""all good""

End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ObjectVariableNotSetInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        public void ObjectVariableNotSet_GivenVariantVariable_ReturnsNoResult()
        {
            const string inputCode = @"
Private Sub Workbook_Open()
    
    Dim target As Variant
    target = Range(""A1"")
    
    target.Value = ""all good""

End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ObjectVariableNotSetInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        public void ObjectVariableNotSet_GivenObjectVariableNotSet_ReturnsResult()
        {
            const string inputCode = @"
Private Sub Workbook_Open()
    
    Dim target As Range
    target = Range(""A1"")
    
    target.Value = ""forgot something?""

End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ObjectVariableNotSetInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        public void ObjectVariableNotSet_GivenSetObjectVariable_ReturnsNoResult()
        {
            const string inputCode = @"
Private Sub Workbook_Open()
    
    Dim target As Range
    Set target = Range(""A1"")
    
    target.Value = ""All good""

End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ObjectVariableNotSetInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }
    }
}