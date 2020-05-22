using System.Linq;
using Moq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class ObsoleteGlobalInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void ObsoleteGlobal_ReturnsResult()
        {
            const string inputCode =
                @"Global var1 As Integer";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ObsoleteGlobal_ReturnsResult_MultipleGlobals()
        {
            const string inputCode =
                @"Global var1 As Integer
Global var2 As String";
            Assert.AreEqual(2, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ObsoleteGlobal_DoesNotReturnResult()
        {
            const string inputCode =
                @"Public var1 As Integer";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ObsoleteGlobal_ReturnsResult_SomeConstantsUsed()
        {
            const string inputCode =
                @"Public var1 As Integer
Global var2 As Date";
        Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ObsoleteGlobal_ReturnsNoResult_BracketedExpression()
        {
            const string code = @"
Public Sub DoSomething()
    [A1] = 42
End Sub
";
            var vbe = MockVbeBuilder.BuildFromModules(("Module1", code, ComponentType.StandardModule), new ReferenceLibrary[] { ReferenceLibrary.VBA, ReferenceLibrary.Excel });
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupGet(m => m.ApplicationName).Returns("Excel");
            vbe.Setup(m => m.HostApplication()).Returns(() => mockHost.Object);

            Assert.AreEqual(0, InspectionResults(vbe.Object).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ObsoleteGlobal_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
                @"'@Ignore ObsoleteGlobal
Global var1 As Integer";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            var inspection = new ObsoleteGlobalInspection(null);

            Assert.AreEqual(nameof(ObsoleteGlobalInspection), inspection.Name);
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new ObsoleteGlobalInspection(state);
        }
    }
}
