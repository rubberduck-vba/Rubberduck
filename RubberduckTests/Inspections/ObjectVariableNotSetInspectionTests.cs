using System.Linq;
using System.Threading;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Common;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    [DeploymentItem(@"TestFiles\")]
    public class ObjectVariableNotSetInspectionTests
    {
        [Test]
        [Category("Inspections")]
        public void ObjectVariableNotSet_AlsoAssignedToNothing_ReturnsNoResult()
        {
            var expectResultCount = 0;
            var input =
                @"
Private Sub DoSomething()
    Dim target As Object
    Set target = New Object
    target.DoSomething
    Set target = Nothing
End Sub
";
            AssertInputCodeYieldsExpectedInspectionResultCount(input, expectResultCount);
        }

        [Test]
        [Category("Inspections")]
        public void ObjectVariableNotSet_GivenIndexerObjectAccess_ReturnsNoResult()
        {
            var expectResultCount = 0;
            var input =
@"
Private Sub DoSomething()
    Dim target As Object
    Set target = CreateObject(""Scripting.Dictionary"")
    target(""foo"") = 42
End Sub
";
            AssertInputCodeYieldsExpectedInspectionResultCount(input, expectResultCount);
        }

        [Test]
        [Category("Inspections")]
        public void ObjectVariableNotSet_GivenPropertyLet_ReturnsNoResult()
        {
            var expectResultCount = 0;
            var input =
                @"
Public Property Let Foo(rhs As String)
End Property

Private Sub DoSomething()
    Foo = 42
End Sub
";
            AssertInputCodeYieldsExpectedInspectionResultCount(input, expectResultCount);
        }

        [Test]
        [Category("Inspections")]
        public void ObjectVariableNotSet_GivenPropertySet_WithoutSet_ReturnsResult()
        {
            var expectResultCount = 1;
            var input =
                @"
Public Property Set Foo(rhs As Object)
End Property

Private Sub DoSomething()
    Foo = New Object
End Sub
";
            AssertInputCodeYieldsExpectedInspectionResultCount(input, expectResultCount);
        }

        [Test]
        [Category("Inspections")]
        public void ObjectVariableNotSet_GivenPropertySet_WithSet_ReturnsNoResult()
        {
            var expectResultCount = 0;
            var input =
                @"
Public Property Set Foo(rhs As Object)
End Property

Private Sub DoSomething()
    Set Foo = New Object
End Sub
";
            AssertInputCodeYieldsExpectedInspectionResultCount(input, expectResultCount);
        }

        [Test]
        [Category("Inspections")]
        public void ObjectVariableNotSet_GivenIndexerObjectAccess_ReturnsResult()
        {
            var expectResultCount = 1;
            var input =
@"
Private Sub DoSomething()
    Dim target As Object
    target = CreateObject(""Scripting.Dictionary"")
    target(""foo"") = 42
End Sub
";
            AssertInputCodeYieldsExpectedInspectionResultCount(input, expectResultCount);
        }

        [Test]
        [Category("Inspections")]
        public void ObjectVariableNotSet_GivenStringVariable_ReturnsNoResult()
        {
            var expectResultCount = 0;
            var input =
@"
Private Sub Workbook_Open()    
    Dim target As String
    target = Range(""A1"")
    target.Value = ""all good""
End Sub";
            AssertInputCodeYieldsExpectedInspectionResultCount(input, expectResultCount, "Excel.1.8.xml");
        }

        [Test]
        [Category("Inspections")]
        public void ObjectVariableNotSet_GivenVariantVariableAssignedObject_ReturnsResult()
        {
            var expectResultCount = 1;
            var input =
@"
Private Sub TestSub(ByRef testParam As Variant)
    Dim target As Collection
    Set target = New Collection
    testParam = target
    testParam.Add 100
End Sub";
            AssertInputCodeYieldsExpectedInspectionResultCount(input, expectResultCount, "VBA.4.2.xml");
        }

        [Test]
        [Category("Inspections")]
        public void ObjectVariableNotSet_GivenVariantVariableAssignedNewObject_ReturnsResult()
        {
            var expectResultCount = 1;
            var input =
@"
Private Sub TestSub(ByRef testParam As Variant)
    testParam = New Collection     
End Sub";
            AssertInputCodeYieldsExpectedInspectionResultCount(input, expectResultCount, "VBA.4.2.xml");
        }

        [Test]
        [Category("Inspections")]
        public void ObjectVariableNotSet_GivenVariantVariableAssignedRange_ReturnsResult()
        {
            var expectResultCount = 1;
            var input =
@"
Private Sub TestSub(ByRef testParam As Variant)
    testParam = Range(""A1:C1"")    
End Sub";
            AssertInputCodeYieldsExpectedInspectionResultCount(input, expectResultCount, "Excel.1.8.xml");
        }

        [Test]
        [Category("Inspections")]
        public void ObjectVariableNotSet_GivenVariantVariableAssignedDeclaredRange_ReturnsResult()
        {
            var expectResultCount = 1;
            var input =
@"
Private Sub TestSub(ByRef testParam As Variant, target As Range)
    testParam = target
End Sub";
            AssertInputCodeYieldsExpectedInspectionResultCount(input, expectResultCount, "Excel.1.8.xml");
        }

        [Test]
        [Category("Inspections")]
        public void ObjectVariableNotSet_GivenVariantVariableAssignedBaseType_ReturnsNoResult()
        {
            var expectResultCount = 0;
            var input =
@"
Private Sub Workbook_Open()
    Dim target As Variant
    target = ""A1""
End Sub";
            AssertInputCodeYieldsExpectedInspectionResultCount(input, expectResultCount);
        }

        [Test]
        [Category("Inspections")]
        public void ObjectVariableNotSet_GivenObjectVariableNotSet_ReturnsResult()
        {
            var expectResultCount = 1;
            var input =
@"
Private Sub Workbook_Open()
    
    Dim target As Range
    target = Range(""A1"")
    
    target.Value = ""forgot something?""

End Sub";
            AssertInputCodeYieldsExpectedInspectionResultCount(input, expectResultCount, "Excel.1.8.xml");
        }

        [Test]
        [Category("Inspections")]
        public void ObjectVariableNotSet_GivenObjectVariableNotSet_Ignored_DoesNotReturnResult()
        {
            var expectResultCount = 0;
            var input =
@"
Private Sub Workbook_Open()
    
    Dim target As Range
'@Ignore ObjectVariableNotSet
    target = Range(""A1"")
    
    target.Value = ""forgot something?""

End Sub";
            AssertInputCodeYieldsExpectedInspectionResultCount(input, expectResultCount, "Excel.1.8.xml");
        }

        [Test]
        [Category("Inspections")]
        public void ObjectVariableNotSet_GivenSetObjectVariable_ReturnsNoResult()
        {
            var expectResultCount = 0;
            var input =
@"
Private Sub Workbook_Open()
    
    Dim target As Range
    Set target = Range(""A1"")
    
    target.Value = ""All good""

End Sub";
            AssertInputCodeYieldsExpectedInspectionResultCount(input, expectResultCount, "Excel.1.8.xml");
        }

        [Test]
        [Category("Inspections")]
        public void ObjectVariableNotSet_LongPtrVariable_ReturnsNoResult()
        {
            var expectResultCount = 0;
            var input =
@"
Private Sub TestLongPtr()
    Dim handle as LongPtr
    handle = 123456
End Sub";
            AssertInputCodeYieldsExpectedInspectionResultCount(input, expectResultCount);
        }

        [Test]
        [Category("Inspections")]
        public void ObjectVariableNotSet_NoTypeSpecified_ReturnsResult()
        {
            var expectResultCount = 0;
            var input =
@"
Private Sub TestLongPtr()
    Dim handle as LongPtr
    handle = 123456
End Sub";
            AssertInputCodeYieldsExpectedInspectionResultCount(input, expectResultCount);
        }

        [Test]
        [Category("Inspections")]
        public void ObjectVariableNotSet_SelfAssigned_ReturnsNoResult()
        {
            var expectResultCount = 0;
            var input =
@"
Private Sub TestSelfAssigned()
    Dim arg1 As new Collection
    arg1.Add 7
End Sub";
            AssertInputCodeYieldsExpectedInspectionResultCount(input, expectResultCount, "VBA.4.2.xml");
        }

        [Test]
        [Category("Inspections")]
        public void ObjectVariableNotSet_UDT_ReturnsNoResult()
        {

            var expectResultCount = 0;
            var input =
@"
Private Type TTest
    Foo As Long
    Bar As String
End Type

Private Sub TestUDT()
    Dim tt As TTest
    tt.Foo = 42
End Sub";
            AssertInputCodeYieldsExpectedInspectionResultCount(input, expectResultCount);
        }

        [Test]
        [Category("Inspections")]
        public void ObjectVariableNotSet_EnumVariable_ReturnsNoResult()
        {

            var expectResultCount = 0;
            var input =
                @"
Enum TestEnum
    EnumOne
    EnumTwo
    EnumThree
End Enum

Private Sub TestEnum()
    Dim enumVariable As TestEnum
    enumVariable = EnumThree
End Sub";
            AssertInputCodeYieldsExpectedInspectionResultCount(input, expectResultCount);
        }

        [Test]
        [Category("Inspections")]
        public void ObjectVariableNotSet_FunctionReturnNotSet_ReturnsResult()
        {
            var expectResultCount = 1;
            var input =
@"
Private Function Test() As Collection
    Test = New Collection
End Function";
            AssertInputCodeYieldsExpectedInspectionResultCount(input, expectResultCount, "VBA.4.2.xml");
        }

        [Test]
        [Category("Inspections")]
        public void ObjectVariableNotSet_ObjectLiteral_ReturnsResult()
        {

            var expectResultCount = 1;
            var input =
    @"
Private Sub Test()
    Dim bar As Variant
    bar = Nothing
End Sub";
            AssertInputCodeYieldsExpectedInspectionResultCount(input, expectResultCount);
        }

        [Test]
        [Category("Inspections")]
        public void ObjectVariableNotSet_NonObjectLiteral_ReturnsNoResult()
        {

            var expectResultCount = 0;
            var input =
    @"
Private Sub Test()
    Dim bar As Variant
    bar = Null
    bar = Empty
    bar = ""aaa""
    bar = 5
End Sub";
            AssertInputCodeYieldsExpectedInspectionResultCount(input, expectResultCount);
        }

        [Test]
        [Category("Inspections")]
        public void ObjectVariableNotSet_ForEach_ReturnsNoResult()
        {

            var expectResultCount = 0;
            var input =
    @"
Private Sub Test()
    Dim bar As Variant
    For Each foo In bar
    Next
End Sub";
            AssertInputCodeYieldsExpectedInspectionResultCount(input, expectResultCount);
        }

        [Test]
        [Category("Inspections")]
        public void ObjectVariableNotSet_RSet_ReturnsNoResult()
        {

            var expectResultCount = 0;
            var input =
    @"
Private Sub Test()
    Dim foo As Variant
    Dim bar As Variant
    RSet foo = bar
End Sub";
            AssertInputCodeYieldsExpectedInspectionResultCount(input, expectResultCount);
        }

        [Test]
        [Category("Inspections")]
        public void ObjectVariableNotSet_LSet_ReturnsNoResult()
        {

            var expectResultCount = 0;
            var input =
    @"
Private Sub Test()
    Dim foo As Variant
    Dim bar As Variant
    LSet foo = bar
End Sub";
            AssertInputCodeYieldsExpectedInspectionResultCount(input, expectResultCount);
        }

        [Test]
        [Category("Inspections")]
        public void ObjectVariableNotSet_LSetOnUDT_ReturnsNoResult()
        {

            var expectResultCount = 0;
            var input =
                @"
Type TFoo
  CountryCode As String * 2
  SecurityNumber As String * 8
End Type

Type TBar
  ISIN As String * 10
End Type

Sub Test()

  Dim foo As TFoo
  Dim bar As TBar

  bar.ISIN = ""DE12345678""
  LSet foo = bar
End Sub";
            AssertInputCodeYieldsExpectedInspectionResultCount(input, expectResultCount);
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = "ObjectVariableNotSetInspection";
            var inspection = new ObjectVariableNotSetInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }

        private void AssertInputCodeYieldsExpectedInspectionResultCount(string inputCode, int expected, params string[] testLibraries)
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", "TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, inputCode)
                .AddReference("VBA", MockVbeBuilder.LibraryPathVBA, 4, 2, true)
                .AddReference("Excel", MockVbeBuilder.LibraryPathMsExcel, 1, 8, true)
                .Build();
            var vbe = builder.AddProject(project).Build();

            using(var coordinator = MockParser.Create(vbe.Object))
            {
                foreach (var testLibrary in testLibraries)
                {
                    coordinator.State.AddTestLibrary(testLibrary);
                }
                coordinator.Parse(new CancellationTokenSource());

                var inspection = new ObjectVariableNotSetInspection(coordinator.State);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(expected, inspectionResults.Count());
            }
        }
    }
}

