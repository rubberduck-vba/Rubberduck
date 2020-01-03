using System.Linq;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class SuspiciousLetAssignmentInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        [TestCase("Class1", "Class2")]
        [TestCase("Object", "Class2")]
        [TestCase("Class1", "Object")]
        [TestCase("Object", "Object")]
        public void BothSidesOfAssignmentHaveDefaultMemberAccess_NoExplicitLet_OneResult(string assignedTypeName, string assignedToTypeName)
        {
            var class1Code = @"
Public Function Foo() As Long
Attribute Foo.VB_UserMemId = 0
End Function
";

            var class2Code = @"
Public Property Let Baz(RHS As Long)
Attribute Baz.VB_UserMemId = 0
End Property
";

            var moduleCode = $@"
Private Sub Bar() 
    Dim cls1 As {assignedTypeName}
    Dim cls2 As {assignedToTypeName} 
    Set cls1 = New Class1
    Set cls2 = New Class2
    cls2 = cls1
End Sub
";

            var inspectionResults = InspectionResultsForModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Class2", class2Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            var inspectionResult = inspectionResults.Single();

            Assert.IsNotNull(inspectionResult.Properties.RhSReference);

            if (assignedTypeName.Equals("Object") || assignedToTypeName.Equals("Object"))
            {
                var deactivatedFixes = inspectionResult.Properties.DisableFixes;
                Assert.AreEqual("ExpandDefaultMemberQuickFix", deactivatedFixes);
            }
        }

        [Test]
        [Category("Inspections")]
        [TestCase("Class1", "Class2")]
        [TestCase("Object", "Class2")]
        [TestCase("Class1", "Object")]
        [TestCase("Object", "Object")]
        public void BothSidesOfAssignmentHaveDefaultMemberAccess_ExplicitLet_NoResult(string assignedTypeName, string assignedToTypeName)
        {
            var class1Code = @"
Public Function Foo() As Long
Attribute Foo.VB_UserMemId = 0
End Function
";

            var class2Code = @"
Public Property Let Baz(RHS As Long)
Attribute Baz.VB_UserMemId = 0
End Property
";

            var moduleCode = $@"
Private Sub Bar() 
    Dim cls1 As {assignedTypeName}
    Dim cls2 As {assignedToTypeName} 
    Set cls1 = New Class1
    Set cls2 = New Class2
    Let cls2 = cls1
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
        //This is covered by ObjectVariableNotSetInspection
        public void LeftSideFailedDefaultMemberResolution_NoResult()
        {
            var class1Code = @"
Public Function Foo() As Long
Attribute Foo.VB_UserMemId = 0
End Function
";

            var class2Code = @"
Public Property Let Baz(RHS As Long)
End Property
";

            var moduleCode = @"
Private Sub Bar() 
    Dim cls1 As Class1
    Dim cls2 As Class2
    Set cls1 = New Class1
    Set cls2 = New Class2
    Let cls2 = cls1
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
        //This is covered by ObjectVariableNotSetInspection
        public void RightSideFailedDefaultMemberResolution_NoResult()
        {
            var class1Code = @"
Public Function Foo() As Long
End Function
";

            var class2Code = @"
Public Property Let Baz(RHS As Long)
Attribute Baz.VB_UserMemId = 0
End Property
";

            var moduleCode = @"
Private Sub Bar() 
    Dim cls1 As Class1
    Dim cls2 As Class2
    Set cls1 = New Class1
    Set cls2 = New Class2
    Let cls2 = cls1
End Sub
";

            var inspectionResults = InspectionResultsForModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Class2", class2Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            Assert.AreEqual(0, inspectionResults.Count());
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new SuspiciousLetAssignmentInspection(state);
        }
    }
}