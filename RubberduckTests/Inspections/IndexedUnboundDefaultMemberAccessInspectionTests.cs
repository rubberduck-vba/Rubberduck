using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class IndexedUnboundDefaultMemberAccessInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void OrdinaryIndexedDefaultMemberAccessExpression_NoResult()
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

            var vbe = MockVbeBuilder.BuildFromModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Class2", class2Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            var inspectionResults = InspectionResults(vbe.Object);

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void UnboundIndexedDefaultMemberAccessExpression_OneResult()
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
    Dim cls As Object
    Set Foo = cls(""newClassObject"")
End Function
";

            var vbe = MockVbeBuilder.BuildFromModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Class2", class2Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            var inspectionResults = InspectionResults(vbe.Object);

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        //The second call in an index access to on unbound expression, which is not identified as default member access.
        public void DoubleUnboundIndexedDefaultMemberAccessExpression_OneResult()
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
    Dim cls As Object
    Set Foo = cls(""newClassObject"")(""whatever"")
End Function
";

            var vbe = MockVbeBuilder.BuildFromModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Class2", class2Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            var inspectionResults = InspectionResults(vbe.Object);

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void RecursiveIndexedDefaultMemberAccessExpression_NoResult()
        {
            var class1Code = @"
Public Function Foo() As Class2
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

            var vbe = MockVbeBuilder.BuildFromModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Class2", class2Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            var inspectionResults = InspectionResults(vbe.Object);

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void UnboundIndexedDefaultMemberAccessOnDefaultMemberArrayAccess_OneResult()
        {
            var class1Code = @"
Public Function Foo() As Object()
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
    Set Foo = cls(0)(""newClassObject"")
End Function
";

            var vbe = MockVbeBuilder.BuildFromModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Class2", class2Code, ComponentType.ClassModule),
                ("Class3", class3Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            var inspectionResults = InspectionResults(vbe.Object);

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void JaggedArray_NoResult()
        {
            var moduleCode = @"
Private Function Foo() As Variant 
    Dim outerArray() As Variant
    Dim firstInnerArray() As Variant
    Dim secondInnerArray() As Variant
    outerArray = Array(firstInnerArray, secondInnerArray)
    Foo = outerArray(0)(0)
End Function
";

            var inspectionResults = InspectionResultsForStandardModule(moduleCode);

            Assert.AreEqual(0, inspectionResults.Count());
        }


        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new IndexedUnboundDefaultMemberAccessInspection(state);
        }
    }
}