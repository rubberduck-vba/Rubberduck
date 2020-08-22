using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class AssignedByValParameterInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void AssignedByValParameter_ReturnsResult_Sub()
        {
            const string inputCode =
@"Public Sub Foo(ByVal arg1 As String)
    Let arg1 = ""test""
End Sub";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void AssignedByValParameter_ReturnsResult_Function()
        {
            const string inputCode =
@"Function Foo(ByVal arg1 As Integer) As Boolean
    Let arg1 = 9
End Function";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void AssignedByValParameter_ReturnsResult_MultipleParams()
        {
            const string inputCode =
@"Public Sub Foo(ByVal arg1 As String, ByVal arg2 As Integer)
    Let arg1 = ""test""
    Let arg2 = 9
End Sub";
            Assert.AreEqual(2, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void AssignedByValParameter_DoesNotReturnResult()
        {
            const string inputCode =
@"Public Sub Foo(ByVal arg1 As String)
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void AssignedByValParameter_Ignored_DoesNotReturnResult_Sub()
        {
            const string inputCode =
@"'@Ignore AssignedByValParameter
Public Sub Foo(ByVal arg1 As String)
    Let arg1 = ""test""
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void AssignedByValParameter_ReturnsResult_SomeAssignedByValParams()
        {
            const string inputCode =
 @"Public Sub Foo(ByVal arg1 As String, ByVal arg2 As Integer)
    Let arg1 = ""test""
    
    Dim var1 As Integer
    var1 = arg2
End Sub";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void AssignedByValParameter_NoResultForLeftHandSideMemberAssignment()
        {
            var class1 = @"
Option Explicit
Private mSomething As Long
Public Property Get Something() As Long
    Something = mSomething
End Property
Public Property Let Something(ByVal value As Long)
    mSomething = value
End Property
";
            var caller = @"
Option Explicit
Private Sub DoSomething(ByVal foo As Class1)
    foo.Something = 42
End Sub
";
            var modules = new(string, string, ComponentType)[] 
            {
                ("Class1", class1, ComponentType.ClassModule),
                ("Module1", caller, ComponentType.StandardModule)
            };

            Assert.AreEqual(0, InspectionResultsForModules(modules).Count());
        }

        [Test]
        [Category("Inspections")]
        public void AssignedByValParameter_NoResultForDefaultMemberAssignment()
        {
            var class1 = @"
Public Property Get Something() As Long
Attribute Something.VB_UserMemId = 0
End Property

Public Property Let Something(ByVal value As Long)
End Property
";
            var caller = @"
Option Explicit
Private Sub DoSomething(ByVal foo As Class1)
    foo = 42
End Sub
";
            var modules = new(string, string, ComponentType)[]
            {
                ("Class1", class1, ComponentType.ClassModule),
                ("Module1", caller, ComponentType.StandardModule)
            };

            Assert.AreEqual(0, InspectionResultsForModules(modules).Count());
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = "AssignedByValParameterInspection";
            var inspection = new AssignedByValParameterInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new AssignedByValParameterInspection(state);
        }
    }
}
