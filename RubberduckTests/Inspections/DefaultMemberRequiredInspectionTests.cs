using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class DefaultMemberRequiredInspectionTests : InspectionTestsBase
    {
        [Category("Inspections")]
        [Test]
        public void ChainedDictionaryAccessFailedAtEnd_OneResult()
        {
            var class1Code = @"
Public Function Foo(bar As String) As Class2
Attribute Foo.VB_UserMemId = 0
    Set Foo = New Class2
End Function
";

            var class2Code = @"
Public Function Baz(bar As String) As Class2
    Set Baz = New Class2
End Function
";

            var moduleCode = @"
Private Function Foo() As Class1 
    Dim cls As new Class1
    Set Foo = cls!newClassObject!whatever
End Function
";

            var vbe = MockVbeBuilder.BuildFromModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Class2", class2Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            var inspectionResults = InspectionResults(vbe.Object);
            var result = inspectionResults.Single();

            var expectedSelection = new Selection(4, 15, 4, 42);
            var actualSelection = result.Context.GetSelection();

            Assert.AreEqual(expectedSelection, actualSelection);
        }

        [Category("Inspections")]
        [Test]
        public void ChainedDictionaryAccessExpressionFailedAtStart_OneResult()
        {
            var class1Code = @"
Public Function Foo(bar As String) As Class2
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
Private Function Foo() As Class1 
    Dim cls As new Class1
    Set Foo = cls!newClassObject!whatever
End Function
";

            var vbe = MockVbeBuilder.BuildFromModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Class2", class2Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            var inspectionResults = InspectionResults(vbe.Object);
            var result = inspectionResults.Single();

            var expectedSelection = new Selection(4, 15, 4, 33);
            var actualSelection = result.Context.GetSelection();

            Assert.AreEqual(expectedSelection, actualSelection);
        }

        [Category("Inspections")]
        [Test]
        public void FailedDictionaryAccessExpressionWithIndexedDefaultMemberAccess_OneResult()
        {
            var class1Code = @"
Public Function Foo(bar As String) As Class2
Attribute Foo.VB_UserMemId = 0
    Set Foo = New Class2
End Function
";

            var class2Code = @"
Public Function Baz(bar As String) As Class2
    Set Baz = New Class2
End Function
";

            var moduleCode = @"
Private Function Foo() As Class1 
    Dim cls As new Class1
    Set Foo = cls!newClassObject(""whatever"")
End Function
";

            var vbe = MockVbeBuilder.BuildFromModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Class2", class2Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            var inspectionResults = InspectionResults(vbe.Object);
            var result = inspectionResults.Single();

            var expectedSelection = new Selection(4, 15, 4, 45);
            var actualSelection = result.Context.GetSelection();

            Assert.AreEqual(expectedSelection, actualSelection);
        }

        [Category("Inspections")]
        [Test]
        public void FailedIndexExpressionOnVariable_WithoutArguments_OneResult()
        {
            var classCode = @"
Public Function Foo() As String
End Function
";

            var moduleCode = @"
Private Function Foo() As String 
    Dim cls As new Class1
    Foo = cls()
End Function
";

            var vbe = MockVbeBuilder.BuildFromModules(
                ("Class1", classCode, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            var inspectionResults = InspectionResults(vbe.Object);
            var result = inspectionResults.Single();

            var expectedSelection = new Selection(4, 11, 4, 16);
            var actualSelection = result.Context.GetSelection();

            Assert.AreEqual(expectedSelection, actualSelection);
        }

        [Category("Inspections")]
        [Test]
        public void FailedIndexExpressionOnVariable_WithArguments_OneResult()
        {
            var classCode = @"
Public Function Foo(index As Long) As String
End Function
";

            var moduleCode = @"
Private Function Foo() As String 
    Dim cls As new Class1
    Foo = cls(0)
End Function
";

            var vbe = MockVbeBuilder.BuildFromModules(
                ("Class1", classCode, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            var inspectionResults = InspectionResults(vbe.Object);
            var result = inspectionResults.Single();

            var expectedSelection = new Selection(4, 11, 4, 17);
            var actualSelection = result.Context.GetSelection();

            Assert.AreEqual(expectedSelection, actualSelection);
        }

        [Category("Inspections")]
        [Test]
        public void FailedIndexExpressionOnArrayAccess_WithoutArguments_OneResult()
        {
            var classCode = @"
Public Function Foo() As String
End Function
";

            var moduleCode = @"
Private Function Foo() As String 
    Dim cls() As new Class1
    Foo = cls(0)()
End Function
";

            var vbe = MockVbeBuilder.BuildFromModules(
                ("Class1", classCode, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            var inspectionResults = InspectionResults(vbe.Object);
            var result = inspectionResults.Single();

            var expectedSelection = new Selection(4, 11, 4, 19);
            var actualSelection = result.Context.GetSelection();

            Assert.AreEqual(expectedSelection, actualSelection);
        }

        [Category("Inspections")]
        [Test]
        public void FailedIndexExpressionOnArrayAccess_WithArguments_OneResult()
        {
            var classCode = @"
Public Function Foo(index As Long) As String
End Function
";

            var moduleCode = @"
Private Function Foo() As String 
    Dim cls() As new Class1
    Foo = cls(0)(2)
End Function
";

            var vbe = MockVbeBuilder.BuildFromModules(
                ("Class1", classCode, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            var inspectionResults = InspectionResults(vbe.Object);
            var result = inspectionResults.Single();

            var expectedSelection = new Selection(4, 11, 4, 20);
            var actualSelection = result.Context.GetSelection();

            Assert.AreEqual(expectedSelection, actualSelection);
        }

        [Category("Inspections")]
        [Test]
        public void FailedIndexExpressionOnOtherIndexExpression_WithoutArguments_OneResult()
        {
            var classCode = @"
Public Function Foo(index As Long) As Class1
End Function
";

            var moduleCode = @"
Private Function Foo() As String 
    Dim cls As new Class1
    Foo = cls.Foo(0)()
End Function
";

            var vbe = MockVbeBuilder.BuildFromModules(
                ("Class1", classCode, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            var inspectionResults = InspectionResults(vbe.Object);
            var result = inspectionResults.Single();

            var expectedSelection = new Selection(4, 11, 4, 23);
            var actualSelection = result.Context.GetSelection();

            Assert.AreEqual(expectedSelection, actualSelection);
        }

        [Category("Inspections")]
        [Test]
        public void FailedIndexExpressionOnOtherIndexExpression_WithArguments_OneResult()
        {
            var classCode = @"
Public Function Foo(index As Long) As Class1
End Function
";

            var moduleCode = @"
Private Function Foo() As String 
    Dim cls As new Class1
    Foo = cls.Foo(0)(2)
End Function
";

            var vbe = MockVbeBuilder.BuildFromModules(
                ("Class1", classCode, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            var inspectionResults = InspectionResults(vbe.Object);
            var result = inspectionResults.Single();

            var expectedSelection = new Selection(4, 11, 4, 24);
            var actualSelection = result.Context.GetSelection();

            Assert.AreEqual(expectedSelection, actualSelection);
        }

        [Category("Inspections")]
        [Test]
        public void FailedIndexExpressionOnParameterlessFunction_WithArguments_OneResult()
        {
            var classCode = @"
Public Function Foo() As Class1
End Function
";

            var moduleCode = @"
Private Function Foo() As String 
    Dim cls As new Class1
    Foo = cls.Foo(0)
End Function
";

            var vbe = MockVbeBuilder.BuildFromModules(
                ("Class1", classCode, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            var inspectionResults = InspectionResults(vbe.Object);
            var result = inspectionResults.Single();

            var expectedSelection = new Selection(4, 11, 4, 21);
            var actualSelection = result.Context.GetSelection();

            Assert.AreEqual(expectedSelection, actualSelection);
        }

        [Category("Inspections")]
        [Test]
        public void OptionalParenthesesAfterVariantReturningProperty_NoResult()
        {
            var classCode = @"
Public Property Get Foo() As Variant
End Property
";

            var moduleCode = @"
Private Function Bar() As String 
    Dim cls As new Class1
    Bar = cls.Foo()
End Function
";

            var vbe = MockVbeBuilder.BuildFromModules(
                ("Class1", classCode, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            var inspectionResults = InspectionResults(vbe.Object);
            
            Assert.AreEqual(0, inspectionResults.Count());
        }

        [Category("Inspections")]
        [Test]
        public void FailedIndexExpressionOnFunctionWithParameters_NoResult()
        {
            var classCode = @"
Public Function Foo(index As Long) As Class1
End Function
";

            var moduleCode = @"
Private Function Foo() As String 
    Dim cls As new Class1
    Foo = cls.Foo(0, 2)
End Function
";

            var vbe = MockVbeBuilder.BuildFromModules(
                ("Class1", classCode, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            var inspectionResults = InspectionResults(vbe.Object);

            Assert.IsFalse(inspectionResults.Any());
        }

        [Category("Inspections")]
        [Test]
        [TestCase("String", "bar = \"Hello \" & Foo(Nothing)")]
        [TestCase("Class1", "Set Foo = Foo(Nothing)")]
        public void RecursiveFunctionCall_NoResult(string functionReturnTypeName, string statement)
        {
            var classCode = @"
Public Function Foo(index As Variant) As Class1
End Function
";

            var moduleCode = $@"
Private Function Foo(ByVal cls As Class1) As {functionReturnTypeName} 
    If Not(cls Is Nothing) Then
        Dim bar As Variant
        {statement}
    End If
End Function
";

            var vbe = MockVbeBuilder.BuildFromModules(
                ("Class1", classCode, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            var inspectionResults = InspectionResults(vbe.Object);

            Assert.IsFalse(inspectionResults.Any());
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new DefaultMemberRequiredInspection(state);
        }
    }
}