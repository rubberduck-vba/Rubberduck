using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class ObjectWhereProcedureIsRequiredInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void NonParameterizedProcedureCoercion_OneResult()
        {
            var class1Code = @"
Public Sub Foo()
Attribute Foo.VB_UserMemId = 0
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

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void NonParameterizedUnboundProcedureCoercion_OneResultWithExpandDefaultMemberQuickFixDisabled()
        {
            var class1Code = @"
Public Sub Foo()
Attribute Foo.VB_UserMemId = 0
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
    Dim cls As Object
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

            var result = inspectionResults.Single();
            var actuallyDisabledQuickFix = result.DisabledQuickFixes.Single();
            Assert.AreEqual("ExpandDefaultMemberQuickFix", actuallyDisabledQuickFix);
        }
        [Test]
        [Category("Inspections")]
        //This is an indexed default member access and in the corresponding inspection. 
        public void ParameterizedProcedureCoercion_NoResult()
        {
            var class1Code = @"
Public Sub Foo()
Attribute Foo.VB_UserMemId = 0
End Sub
";

            var class2Code = @"
Public Function Baz(arg As Long) As Class1
Attribute Baz.VB_UserMemId = 0
    Set Baz = New Class1
End Function
";

            var moduleCode = $@"
Private Function Foo() As Variant 
    Dim cls As new Class2
    cls 42
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

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        //This is an indexed default member access and in the corresponding inspection. 
        public void ParameterizedUnboundProcedureCoercion_NoResult()
        {
            var class1Code = @"
Public Sub Foo()
Attribute Foo.VB_UserMemId = 0
End Sub
";

            var class2Code = @"
Public Function Baz(arg As Long) As Class1
Attribute Baz.VB_UserMemId = 0
    Set Baz = New Class1
End Function
";

            var moduleCode = $@"
Private Function Foo() As Variant 
    Dim cls As Object
    cls 42
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

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void NonParameterizedProcedureCoercion_ExplicitCall_OneResult()
        {
            var class1Code = @"
Public Sub Foo()
Attribute Foo.VB_UserMemId = 0
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

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void NonParameterizedUnboundProcedureCoercion_ExplicitCall_OneResultWithExpandDefaultMemberQuickFixDisabled()
        {
            var class1Code = @"
Public Sub Foo()
Attribute Foo.VB_UserMemId = 0
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
    Dim cls As Object
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

            var result = inspectionResults.Single();
            var actuallyDisabledQuickFix = result.DisabledQuickFixes.Single();
            Assert.AreEqual("ExpandDefaultMemberQuickFix", actuallyDisabledQuickFix);
        }

        [Test]
        [Category("Inspections")]
        //This is an indexed default member access and in the corresponding inspection. 
        public void ParameterizedProcedureCoercion_ExplicitCall_NoResult()
        {
            var class1Code = @"
Public Sub Foo()
Attribute Foo.VB_UserMemId = 0
End Sub
";

            var class2Code = @"
Public Function Baz(arg As Long) As Class1
Attribute Baz.VB_UserMemId = 0
    Set Baz = New Class1
End Function
";

            var moduleCode = $@"
Private Function Foo() As Variant 
    Dim cls As new Class2
    Call cls(42)
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

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        //This is an indexed default member access and in the corresponding inspection. 
        public void ParameterizedUnboundProcedureCoercion_ExplicitCall_NoResult()
        {
            var class1Code = @"
Public Sub Foo()
Attribute Foo.VB_UserMemId = 0
End Sub
";

            var class2Code = @"
Public Function Baz(arg As Long) As Class1
Attribute Baz.VB_UserMemId = 0
    Set Baz = New Class1
End Function
";

            var moduleCode = $@"
Private Function Foo() As Variant 
    Dim cls As Object
    Call cls(42)
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

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void NonParameterizedProcedureCoercionDefaultMemberAccessOnArrayAccess_OneResult()
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

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void NonParameterizedProcedureCoercionDefaultMemberAccessOnArrayAccess_ExplicitCall_OneResult()
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

            Assert.AreEqual(1, inspectionResults.Count());
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new ObjectWhereProcedureIsRequiredInspection(state);
        }
    }
}