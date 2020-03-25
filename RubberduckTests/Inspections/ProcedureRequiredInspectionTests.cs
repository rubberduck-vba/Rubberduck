using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class ProcedureRequiredInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        //This will be handled by another inspection, since it is a failed indexed default member resolution.
        public void FailedParameterizedProcedureCoercionReferenceOnEntireContext()
        {
            var class1Code = @"
Public Sub Foo(arg As Long)
End Sub
";

            var class2Code = @"
Public Function Baz() As Class1
    Set Baz = New Class1
End Function
";

            var moduleCode = $@"
Private Function Foo() As Variant 
    Dim cls As new Class2
    cls.Baz 42
End Function

Private Sub Bar(arg As Long)
End Sub

Private Sub Baz(arg As Variant)
End Sub
";

            var inspectionResults = InspectionResultsForModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Class2", class2Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void FailedNonParameterizedProcedureCoercionReferenceOnEntireContext()
        {
            var class1Code = @"
Public Sub Foo()
End Sub
";

            var class2Code = @"
Public Function Baz() As Class1
    Set Baz = New Class1
End Function
";

            var moduleCode = $@"
Private Function Foo() As Variant 
    Dim cls As new Class2
    cls
End Function

Private Sub Bar(arg As Long)
End Sub

Private Sub Baz(arg As Variant)
End Sub
";

            var inspectionResults = InspectionResultsForModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Class2", class2Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            Assert.AreEqual(1,inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        //This will be handled by another inspection, since it is a failed indexed default member resolution.
        public void FailedParameterizedProcedureCoercionReferenceOnEntireContext_ExplicitCall()
        {
            var class1Code = @"
Public Sub Foo(arg As Long)
End Sub
";

            var class2Code = @"
Public Function Baz() As Class1
    Set Baz = New Class1
End Function
";

            var moduleCode = $@"
Private Function Foo() As Variant 
    Dim cls As new Class2
    Call cls.Baz(42)
End Function

Private Sub Bar(arg As Long)
End Sub

Private Sub Baz(arg As Variant)
End Sub
";

            var inspectionResults = InspectionResultsForModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Class2", class2Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void FailedNonParameterizedProcedureCoercionReferenceOnEntireContext_ExplicitCall()
        {
            var class1Code = @"
Public Sub Foo()
End Sub
";

            var class2Code = @"
Public Function Baz() As Class1
    Set Baz = New Class1
End Function
";

            var moduleCode = $@"
Private Function Foo() As Variant 
    Dim cls As new Class2
    Call cls
End Function

Private Sub Bar(arg As Long)
End Sub

Private Sub Baz(arg As Variant)
End Sub
";

            var inspectionResults = InspectionResultsForModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Class2", class2Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void FailedNonParameterizedProcedureCoercionOnArrayAccessReferenceOnEntireContext()
        {
            var class1Code = @"
Public Sub Foo()
End Sub
";

            var class2Code = @"
Public Function Baz() As Class1()
    Set Baz = New Class1
End Function
";

            var moduleCode = $@"
Private Function Foo() As Variant 
    Dim cls As new Class2
    cls.Baz(42)
End Function

Private Sub Bar(arg As Long)
End Sub

Private Sub Baz(arg As Variant)
End Sub
";

            var inspectionResults = InspectionResultsForModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Class2", class2Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void FailedNonParameterizedProcedureCoercionOnArrayAccessReferenceOnEntireContext_ExplicitCall()
        {
            var class1Code = @"
Public Sub Foo()
End Sub
";

            var class2Code = @"
Public Function Baz() As Class1()
    Set Baz = New Class1
End Function
";

            var moduleCode = $@"
Private Function Foo() As Variant 
    Dim cls As new Class2
    Call cls.Baz(42)
End Function

Private Sub Bar(arg As Long)
End Sub

Private Sub Baz(arg As Variant)
End Sub
";

            var inspectionResults = InspectionResultsForModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Class2", class2Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void SuccessfulParameterizedProcedureCoercionReferenceOnEntireContext()
        {
            var class1Code = @"
Public Sub Foo(arg As Long)
End Sub
";

            var class2Code = @"
Public Function Baz() As Class1
Attribute Baz.VB_UserMemId = 0
    Set Baz = New Class1
End Function
";

            var moduleCode = $@"
Private Function Foo() As Variant 
    Dim cls As new Class2
    cls.Baz 42
End Function

Private Sub Bar(arg As Long)
End Sub

Private Sub Baz(arg As Variant)
End Sub
";

            var inspectionResults = InspectionResultsForModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Class2", class2Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void SuccessfulNonParameterizedProcedureCoercionReferenceOnEntireContext()
        {
            var class1Code = @"
Public Sub Foo()
End Sub
";

            var class2Code = @"
Public Function Baz() As Class1
Attribute Baz.VB_UserMemId = 0
    Set Baz = New Class1
End Function
";

            var moduleCode = $@"
Private Function Foo() As Variant 
    Dim cls As new Class2
    cls
End Function

Private Sub Bar(arg As Long)
End Sub

Private Sub Baz(arg As Variant)
End Sub
";

            var inspectionResults = InspectionResultsForModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Class2", class2Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void SuccessfulParameterizedProcedureCoercionReferenceOnEntireContext_ExplicitCall()
        {
            var class1Code = @"
Public Sub Foo(arg As Long)
End Sub
";

            var class2Code = @"
Public Function Baz() As Class1
Attribute Baz.VB_UserMemId = 0
    Set Baz = New Class1
End Function
";

            var moduleCode = $@"
Private Function Foo() As Variant 
    Dim cls As new Class2
    Call cls.Baz(42)
End Function

Private Sub Bar(arg As Long)
End Sub

Private Sub Baz(arg As Variant)
End Sub
";

            var inspectionResults = InspectionResultsForModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Class2", class2Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void SuccessfulNonParameterizedProcedureCoercionReferenceOnEntireContext_ExplicitCall()
        {
            var class1Code = @"
Public Sub Foo()
End Sub
";

            var class2Code = @"
Public Function Baz() As Class1
Attribute Baz.VB_UserMemId = 0
    Set Baz = New Class1
End Function
";

            var moduleCode = $@"
Private Function Foo() As Variant 
    Dim cls As new Class2
    Call cls
End Function

Private Sub Bar(arg As Long)
End Sub

Private Sub Baz(arg As Variant)
End Sub
";

            var inspectionResults = InspectionResultsForModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Class2", class2Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void SuccessfulNonParameterizedProcedureCoercionOnArrayAccessReferenceOnEntireContext()
        {
            var class1Code = @"
Public Sub Foo()
Attribute Foo.VB_UserMemId = 0
End Sub
";

            var class2Code = @"
Public Function Baz() As Class1()
Attribute Baz.VB_UserMemId = 0
    Set Baz = New Class1
End Function
";

            var moduleCode = $@"
Private Function Foo() As Variant 
    Dim cls As new Class2
    cls.Baz(42)
End Function

Private Sub Bar(arg As Long)
End Sub

Private Sub Baz(arg As Variant)
End Sub
";

            var inspectionResults = InspectionResultsForModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Class2", class2Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void SuccessfulNonParameterizedProcedureCoercionOnArrayAccessReferenceOnEntireContext_ExplicitCall()
        {
            var class1Code = @"
Public Sub Foo()
Attribute Foo.VB_UserMemId = 0
End Sub
";

            var class2Code = @"
Public Function Baz() As Class1()
Attribute Baz.VB_UserMemId = 0
    Set Baz = New Class1
End Function
";

            var moduleCode = $@"
Private Function Foo() As Variant 
    Dim cls As new Class2
    Call cls.Baz(42)
End Function

Private Sub Bar(arg As Long)
End Sub

Private Sub Baz(arg As Variant)
End Sub
";

            var inspectionResults = InspectionResultsForModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Class2", class2Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            Assert.IsFalse(inspectionResults.Any());
        }

        [Category("Inspections")]
        [Test]
        [TestCase("", "Call Foo")]
        [TestCase("", "Call Foo()")]
        [TestCase("", "Foo")]
        [TestCase("ByVal arg As Variant", "Call Foo(arg)")]
        [TestCase("ByVal arg As Variant", "Foo arg")]
        public void RecursiveProcedureCall_NoResult(string argumentList, string statement)
        {
            var classCode = @"
Public Function Foo(index As Variant) As Class1
End Function
";

            var moduleCode = $@"
Private Function Foo({argumentList}) As Class1
    {statement}
End Function
";

            var inspectionResults = InspectionResultsForModules(
                ("Class1", classCode, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

           Assert.IsFalse(inspectionResults.Any());
        }


        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new ProcedureRequiredInspection(state);
        }
    }
}