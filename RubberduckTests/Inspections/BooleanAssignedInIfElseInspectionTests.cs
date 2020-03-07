using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class BooleanAssignedInIfElseInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void Simple()
        {
            const string inputcode =
                @"Sub Foo()
    Dim d As Boolean
    If True Then
        d = True
    Else
        d = False
    EndIf
End Sub";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputcode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void QualifiedName()
        {
            const string inputcode =
                @"Sub Foo()
    If True Then
        Fizz.Buzz = True
    Else
        Fizz.Buzz = False
    EndIf
End Sub";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputcode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void MultipleResults()
        {
            const string inputcode =
                @"Sub Foo()
    Dim d As Boolean
    If True Then
        d = True
    Else
        d = False
    EndIf

    If True Then
        d = False
    Else
        d = True
    EndIf
End Sub";
            Assert.AreEqual(2, InspectionResultsForStandardModule(inputcode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void AssignsInteger()
        {
            const string inputcode =
                @"Sub Foo()
    Dim d
    If True Then
        d = 0
    Else
        d = 1
    EndIf
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputcode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void AssignsTheSameValue()       // worthy of an inspection in its own right
        {
            const string inputcode =
                @"Sub Foo()
    Dim d
    If True Then
        d = True
    Else
        d = True
    EndIf
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputcode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void AssignsToDifferentVariables()
        {
            const string inputcode =
                @"Sub Foo()
    Dim d1, d2
    If True Then
        d1 = True
    Else
        d2 = False
    EndIf
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputcode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ConditionalContainsElseIfBlock()
        {
            const string inputcode =
                @"Sub Foo()
    Dim d
    If True Then
        d = True
    ElseIf False Then
        d = True
    Else
        d = False
    EndIf
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputcode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ConditionalDoesNotContainElseBlock()
        {
            const string inputcode =
                @"Sub Foo()
    Dim d
    If True Then
        d = True
    EndIf
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputcode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void IsIgnored()
        {
            const string inputcode =
                @"Sub Foo()
    Dim d
    '@Ignore BooleanAssignedInIfElse
    If True Then
        d = True
    Else
        d = False
    EndIf
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputcode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void BlockContainsPrefixComment()
        {
            const string inputcode =
                @"Sub Foo()
    Dim d
    If True Then
        ' test
        d = True
    Else
        d = False
    EndIf
End Sub";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputcode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void BlockContainsPostfixComment()
        {
            const string inputcode =
                @"Sub Foo()
    Dim d
    If True Then
        d = True
        ' test
    Else
        d = False
    EndIf
End Sub";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputcode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void BlockContainsEOLComment()
        {
            const string inputcode =
                @"Sub Foo()
    Dim d
    If True Then
        d = True    ' test
    Else
        d = False
    EndIf
End Sub";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputcode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = "BooleanAssignedInIfElseInspection";
            var inspection = new BooleanAssignedInIfElseInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new BooleanAssignedInIfElseInspection(state);
        }
    }
}