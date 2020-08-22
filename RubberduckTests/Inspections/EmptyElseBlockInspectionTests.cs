using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class EmptyElseBlockInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            var inspection = new EmptyElseBlockInspection(null);

            Assert.AreEqual(nameof(EmptyElseBlockInspection), inspection.Name);
        }

        [Test]
        [Category("Inspections")]
        public void EmptyElseBlock_DoesntFireOnEmptyIfBlock()
        {
            const string inputcode =
                @"Sub Foo()
    If True Then
    EndIf
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputcode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void EmptyElseBlock_HasNoContent()
        {
            const string inputcode =
@"Sub Foo()
    If True Then
    Else
    End If
End Sub";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputcode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void EmptyElseBlock_HasQuoteComment()
        {
            const string inputCode =
                @"Sub Foo()
    If True Then
    Else
        'Some Comment
    End If
End Sub";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void EmptyElseBlock_HasRemComment()
        {
            const string inputCode =
                @"Sub Foo()
    If True Then
    Else
        Rem Some Comment
    End If
End Sub";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void EmptyElseBlock_HasVariableDeclaration()
        {
            const string inputCode =
                @"Sub Foo()
    If True Then
    Else
        Dim d
    End If
End Sub";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void EmptyElseBlock_HasConstDeclaration()
        {
            const string inputCode =
                @"Sub Foo()
    If True Then
    Else
        Const c = 0
    End If
End Sub";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void EmptyElseBlock_HasWhitespace()
        {
            const string inputcode =
                @"Sub Foo()
    If True Then
    Else
    
    
    End If
End Sub";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputcode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void EmptyElseBlock_HasDeclarationStatement()
        {
            const string inputCode =
                @"Sub Foo()
    If True Then
    Else
        Dim d
    End If
End Sub";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void EmptyElseBlock_HasExecutableStatement()
        {
            const string inputCode =
                @"Sub Foo()
    If True Then
    Else
        Dim d
        d = 0
    End If
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new EmptyElseBlockInspection(state);
        }
    }
}
