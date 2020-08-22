using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class EmptyIfBlockInspectionTests : InspectionTestsBase
    {
        [TestCase("", 1)]
        [TestCase("ElseIf False Then", 2)]
        [TestCase("Else", 1)]
        [TestCase("' Im a comment", 1)]
        [TestCase("Rem Im a comment", 1)]
        [TestCase("Dim d", 1)]
        [TestCase("Const c = 0", 1)]
        [Category("Inspections")]
        public void EmptyIfBlock_FiresOnSimpleEmptyBlockScenarios(string ifBlockContent, int expected)
        {
            string inputCode =
                $@"Sub Foo()
    If True Then
        {ifBlockContent}
    End If
End Sub";
            Assert.AreEqual(expected, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void EmptyIfBlock_FiresOnEmptySingleLineIfStmt()
        {
            const string inputCode =
                @"Sub Foo()
    If True Then Else Bar
End Sub

Sub Bar()
End Sub";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void EmptyIfBlock_FiresOnEmptyIfBlock_HasNonEmptyElseBlock()
        {
            const string inputCode =
                @"Sub Foo()
    If True Then
    Else
        Dim d
        d = 0
    End If
End Sub";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void EmptyIfBlock_FiresOnEmptyIfBlock_HasWhitespace()
        {
            const string inputCode =
                @"Sub Foo()
    If True Then

    	
    End If
End Sub";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void EmptyIfBlock_IfBlockHasExecutableStatement()
        {
            const string inputCode =
                @"Sub Foo()
    If True Then
        Dim d
        d = 0
    End If
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void EmptyIfBlock_SingleLineIfBlockHasExecutableStatement()
        {
            const string inputCode =
                @"Sub Foo()
    If True Then Bar
End Sub

Sub Bar()
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void EmptyIfBlock_IfAndElseIfBlockHaveExecutableStatement()
        {
            const string inputCode =
                @"Sub Foo()
    If True Then
        Dim d
        d = 0
    ElseIf False Then
        Dim b
        b = 0
    End If
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            var inspection = new EmptyIfBlockInspection(null);

            Assert.AreEqual(nameof(EmptyIfBlockInspection), inspection.Name);
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new EmptyIfBlockInspection(state);
        }
    }
}
