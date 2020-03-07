using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class UseOfBangNotationInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void DictionaryAccessExpression_OneResult()
        {
            var class1Code = @"
Public Function Foo(bar As String) As Class2
Attribute Foo.VB_UserMemId = 0
End Function
";

            var class2Code = @"
Public Function Baz(bar As String) As Class2
Attribute Baz.VB_UserMemId = 0
End Function
";

            var moduleCode = @"
Private Function Foo() As Class2 
    Dim cls As New Class1
    Set Foo = cls!newClassObject
End Function
";

            var inspectionResults = InspectionResultsForModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Class2", class2Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            Assert.AreEqual(1, inspectionResults.Count());
        }
        [Test]
        [Category("Inspections")]
        public void WithDictionaryAccessExpression_OneResult()
        {
            var class1Code = @"
Public Function Foo(bar As String) As Class2
Attribute Foo.VB_UserMemId = 0
End Function
";

            var class2Code = @"
Public Function Baz(bar As String) As Class2
Attribute Baz.VB_UserMemId = 0
End Function
";

            var moduleCode = @"
Private Function Foo() As Class2 
    With New Class1
        Set Foo = !newClassObject
    End With
End Function
";

            var inspectionResults = InspectionResultsForModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Class2", class2Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void ChainedDictionaryAccessExpression_OneResultEach()
        {
            var class1Code = @"
Public Function Foo(bar As String) As Class2
Attribute Foo.VB_UserMemId = 0
End Function
";

            var class2Code = @"
Public Function Baz(bar As String) As Class2
Attribute Baz.VB_UserMemId = 0
End Function
";

            var moduleCode = @"
Private Function Foo() As Class2 
    Dim cls As New Class1
    Set Foo = cls!newClassObject!whatever
End Function
";

            var inspectionResults = InspectionResultsForModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Class2", class2Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

        }

        [Test]
        [Category("Inspections")]
        public void IndexedDefaultMemberAccessExpression_NoResult()
        {
            var class1Code = @"
Public Function Foo(bar As String) As Class2
Attribute Foo.VB_UserMemId = 0
End Function
";

            var class2Code = @"
Public Function Baz(bar As String) As Class2
Attribute Baz.VB_UserMemId = 0
End Function
";

            var moduleCode = @"
Private Function Foo() As Class2 
    Dim cls As New Class1
    Set Foo = cls(""newClassObject"")
End Function
";

            var inspectionResults = InspectionResultsForModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Class2", class2Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void ExplicitCall_NoResult()
        {
            var class1Code = @"
Public Function Foo(bar As String) As Class2
Attribute Foo.VB_UserMemId = 0
End Function
";

            var class2Code = @"
Public Function Baz(bar As String) As Class2
Attribute Baz.VB_UserMemId = 0
End Function
";

            var moduleCode = @"
Private Function Foo() As Class2 
    Dim cls As New Class1
    Set Foo = cls.Foo(""newClassObject"")
End Function
";

            var inspectionResults = InspectionResultsForModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Class2", class2Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        //Covered by UseOfUnboundBangInspection.
        public void UnboundDictionaryAccessExpression_NoResult()
        {
            var class1Code = @"
Public Function Foo(bar As String) As Class2
Attribute Foo.VB_UserMemId = 0
    Set Foo = New Class2
End Function
";

            var class2Code = @"
Public Function Baz(bar As String) As Class2
Attribute Baz.VB_UserMemId = 0
    Set Baz = New Class2
End Function
";

            var moduleCode = @"
Private Function Foo() As Class2 
    Dim cls As Object
    Set Foo = cls!newClassObject
End Function
";

            var inspectionResults = InspectionResultsForModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Class2", class2Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void DictionaryAccessOnDefaultMemberArrayAccess_OneResult()
        {
            var class1Code = @"
Public Function Foo() As Class2()
Attribute Foo.VB_UserMemId = 0
End Function
";

            var class2Code = @"
Public Function Baz(bar As String) As Class2
Attribute Baz.VB_UserMemId = 0
End Function
";

            var moduleCode = @"
Private Function Foo() As Class2 
    Dim cls As New Class1
    Set Foo = cls(0)!newClassObject
End Function
";

            var inspectionResults = InspectionResultsForModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Class2", class2Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        //Covered by UseOfRecursiveBangInspection.
        public void RecursiveDictionaryAccessOnDefaultMemberArrayAccess_NoResult()
        {
            var class1Code = @"
Public Function Foo() As Class3()
Attribute Foo.VB_UserMemId = 0
End Function
";

            var class2Code = @"
Public Function Baz(bar As String) As Class2
Attribute Baz.VB_UserMemId = 0
End Function
";

            var class3Code = @"
Public Function Baz() As Class2
Attribute Baz.VB_UserMemId = 0
End Function
";

            var moduleCode = @"
Private Function Foo() As Class2 
    Dim cls As New Class1
    Set Foo = cls(0)!newClassObject
End Function
";

            var inspectionResults = InspectionResultsForModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Class2", class2Code, ComponentType.ClassModule),
                ("Class3", class3Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        //Covered by UseOfUnboundBangInspection.
        public void RecursiveUnboundDictionaryAccessExpression_NoResult()
        {
            var class1Code = @"
Public Function Foo() As Object
Attribute Foo.VB_UserMemId = 0
End Function
";

            var class2Code = @"
Public Function Baz(bar As String) As Class2
Attribute Baz.VB_UserMemId = 0
End Function
";

            var moduleCode = @"
Private Function Foo() As Class2 
    Dim cls As New Class1
    Set Foo = cls!newClassObject
End Function
";

            var inspectionResults = InspectionResultsForModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Class2", class2Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            Assert.AreEqual(0, inspectionResults.Count());
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new UseOfBangNotationInspection(state);
        }
    }
}