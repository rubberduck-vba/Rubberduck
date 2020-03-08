using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class ObsoleteLetStatementInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void ObsoleteLetStatement_ReturnsResult()
        {
            const string inputCode =
                @"Public Sub Foo()
    Dim var1 As Integer
    Dim var2 As Integer
    
    Let var2 = var1
End Sub";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ObsoleteLetStatement_ReturnsResult_MultipleLets()
        {
            const string inputCode =
                @"Public Sub Foo()
    Dim var1 As Integer
    Dim var2 As Integer
    
    Let var2 = var1
    Let var1 = var2
End Sub";
            Assert.AreEqual(2, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ObsoleteLetStatement_DoesNotReturnResult()
        {
            const string inputCode =
                @"Public Sub Foo()
    Dim var1 As Integer
    Dim var2 As Integer
    
    var2 = var1
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ObsoleteLetStatement_ReturnsResult_SomeConstantsUsed()
        {
            const string inputCode =
                @"Public Sub Foo()
    Dim var1 As Integer
    Dim var2 As Integer
    
    Let var2 = var1
    var1 = var2
End Sub";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ObsoleteLetStatement_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
                @"Public Sub Foo()
    Dim var1 As Integer
    Dim var2 As Integer
    
    '@Ignore ObsoleteLetStatement
    Let var2 = var1
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            var inspection = new ObsoleteLetStatementInspection(null);

            Assert.AreEqual(nameof(ObsoleteLetStatementInspection), inspection.Name);
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new ObsoleteLetStatementInspection(state);
        }
    }
}
