using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class ObsoleteCallStatementInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void ObsoleteCallStatement_ReturnsResult()
        {
            const string inputCode =
                @"Sub Foo()
    Call Foo
End Sub";

            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ObsoleteCallStatement_DoesNotReturnResult()
        {
            const string inputCode =
                @"Sub Foo()
    Foo
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ObsoleteCallStatement_DoesNotReturnResult_InstructionSeparator()
        {
            const string inputCode =
                @"Sub Foo()
    Call Foo: Foo
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ObsoleteCallStatement_ReturnsResult_LabelInFront()
        {
            const string inputCode =
                @"Sub Foo()
    Foo: Call Foo
End Sub";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ObsoleteCallStatement_ReturnsResult_LabelInFrontWithSeparator()
        {
            const string inputCode =
                @"Sub Foo()
    Foo: Call Foo : Foo
End Sub";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ObsoleteCallStatement_ReturnsResult_ColonInComment()
        {
            const string inputCode =
                @"Sub Foo()
    Call Foo ' I''ve got a colon: see?
End Sub";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ObsoleteCallStatement_ReturnsResult_ColonInStringLiteral()
        {
            const string inputCode =
                @"Sub Foo(ByVal str As String)
    Call Foo("":"")
End Sub";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ObsoleteCallStatement_ReturnsMultipleResults()
        {
            const string inputCode =
                @"Sub Foo()
    Call Goo(1, ""test"")
End Sub

Sub Goo(arg1 As Integer, arg1 As String)
    Call Foo
End Sub";
            Assert.AreEqual(2, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ObsoleteCallStatement_ReturnsResults_SomeObsoleteCallStatements()
        {
            const string inputCode =
                @"Sub Foo()
    Call Goo(1, ""test"")
End Sub

Sub Goo(arg1 As Integer, arg1 As String)
    Foo
End Sub";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ObsoleteCallStatement_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
                @"Sub Foo()
    '@Ignore ObsoleteCallStatement
    Call Foo
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            var inspection = new ObsoleteCallStatementInspection(null);

            Assert.AreEqual(nameof(ObsoleteCallStatementInspection), inspection.Name);
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new ObsoleteCallStatementInspection(state);
        }
    }
}
