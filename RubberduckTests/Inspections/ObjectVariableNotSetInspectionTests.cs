using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class ObjectVariableNotSetInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void ObjectVariableNotSet_NotResultForNonObjectPropertyGetWithObjectArgument()
        {
            var expectedResultCount = 0;
            var input = @"
Public Property Get Foo(ByVal bar As Object) As Boolean
    Foo = True
End Property
";
            AssertInputCodeYieldsExpectedInspectionResultCount(input, expectedResultCount);
        }

        [Test]
        [Category("Inspections")]
        public void ObjectVariableNotSet_AlsoAssignedToNothing_ReturnsNoResult()
        {
            var expectResultCount = 0;
            var input =
                @"
Private Sub DoSomething()
    Dim target As Object
    Set target = New Class1
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
    Foo = New Class1
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
    Set Foo = New Class1
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
Public Function Item(index As Variant) As Class1
Attribute Item.VB_UserMemId = 0
End Function

Private Sub DoSomething()
    Dim target As Class1
    Set target = New Class1
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
            AssertInputCodeYieldsExpectedInspectionResultCount(input, expectResultCount, ReferenceLibrary.Excel);
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
            AssertInputCodeYieldsExpectedInspectionResultCount(input, expectResultCount, ReferenceLibrary.VBA);
        }

        [Test]
        [Category("Inspections")]
        //Let assignments to a variable with declared type Object resolve to an unbound default member call.
        //Whether it is legal can only be determined at runtime. However, this creates results for other inspections.
        public void ObjectVariableNotSet_GivenObjectVariableAssignedObject_ReturnsNoResult()
        {
            var expectResultCount = 0;
            var input =
                @"
Public Function Item() As Long
Attribute Item.VB_UserMemId = 0
End Function

Private Sub TestSub(ByRef testParam As Variant)
    Dim target As Object
    Dim foo As Class1
    Set foo = New Class1
    target = foo
End Sub
";
            AssertInputCodeYieldsExpectedInspectionResultCount(input, expectResultCount);
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
            AssertInputCodeYieldsExpectedInspectionResultCount(input, expectResultCount, ReferenceLibrary.VBA);
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
            AssertInputCodeYieldsExpectedInspectionResultCount(input, expectResultCount, ReferenceLibrary.Excel);
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
            AssertInputCodeYieldsExpectedInspectionResultCount(input, expectResultCount, ReferenceLibrary.Excel);
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
            AssertInputCodeYieldsExpectedInspectionResultCount(input, expectResultCount, ReferenceLibrary.VBA);
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

        // This is a corner case similar to #4037. Previously, Collection's default member was not being generated correctly in
        // when it was loaded by the COM collector (_Collection is missing the default interface flag). After picking up that member
        // this test fails because it resolves as attempting to assign 'New Collection' to `Test.DefaultMember`.
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
            AssertInputCodeYieldsExpectedInspectionResultCount(input, expectResultCount, ReferenceLibrary.VBA);
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
        public void ObjectVariableNotSet_ForEachObject_ReturnsNoResult()
        {

            var expectResultCount = 0;
            var input =
                @"
Private Sub Test()
    Dim bar As Object
    For Each foo In bar
    Next
End Sub";
            AssertInputCodeYieldsExpectedInspectionResultCount(input, expectResultCount);
        }

        [Test]
        [Category("Inspections")]
        public void ObjectVariableNotSet_InsideForEachObject_ReturnsResult()
        {

            var expectResultCount = 1;
            var input =
                @"
Private Sub Test()
    Dim bar As Variant
    Dim baz As Class1
    For Each foo In bar
        baz = foo
    Next
End Sub";
            AssertInputCodeYieldsExpectedInspectionResultCount(input, expectResultCount);
        }

        [Test]
        [Category("Inspections")]
        public void ObjectVariableNotSet_InsideForEachSetObject_ReturnsNoResult()
        {

            var expectResultCount = 0;
            var input =
                @"
Private Sub Test()
    Dim bar As Variant
    Dim baz As Object
    For Each foo In bar
        Set baz = foo
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
        public void ObjectVariableNotSet_ComplexExpressionOnRHSWithMemberAccess_ReturnsNoResult()
        {

            var expectResultCount = 0;
            var input =
                @"
Private Sub Test()
    Dim foo As Variant
    Dim bar As Collection
    Set bar = New Collection
    bar.Add ""x"", ""x""
    foo = ""Test"" & bar.Item(""x"")
End Sub";
            AssertInputCodeYieldsExpectedInspectionResultCount(input, expectResultCount, ReferenceLibrary.VBA);
        }

        [Test]
        [Category("Inspections")]
        public void ObjectVariableNotSet_SingleRHSVariableCaseRespectsDeclarationShadowing()
        {

            var expectResultCount = 0;
            var input =
                @"
Private bar As Collection

Private Sub Test()
    Dim foo As Variant
    Dim bar As Long
    bar = 42
    foo = bar
End Sub";
            AssertInputCodeYieldsExpectedInspectionResultCount(input, expectResultCount);
        }

        [Test]
        [Category("Inspections")]
        public void ObjectVariableNotSet_SingleRHSVariableCaseRespectsDefaultMembers()
        {
            var expectResultCount = 0;
            var input =
                @"
Private Sub Test()    
    Dim foo As Range
    Dim bar As Variant    
    bar = foo
End Sub";
            AssertInputCodeYieldsExpectedInspectionResultCount(input, expectResultCount, ReferenceLibrary.Excel);
        }

        [Test]
        [Category("Inspections")]
        public void ObjectVariableNotSet_SingleRHSVariableCaseIdentifiesDefaultMembersNotReturningAnObject()
        {
            var expectResultCount = 1;
            var input =
                @"
Private Sub Test()    
    Dim foo As Recordset
    Dim bar As Variant    
    bar = foo
End Sub";
            //The default member of Recordset is Fields, which is an object.
            AssertInputCodeYieldsExpectedInspectionResultCount(input, expectResultCount, ReferenceLibrary.AdoDb);
        }

        [Test]
        [Category("Inspections")]
        public void ObjectVariableNotSet_AssignmentToVariableWithDefaultMemberReturningAnObjectOfASpecificClassWithoutParameterlessDefaultMember_OneResult()
        {
            var expectResultCount = 1;
            var input =
                @"
Private Sub Test()    
    Dim foo As Recordset
    Dim bar As Variant    
    foo = bar
End Sub";
            //The default member of Recordset is Fields, which is an object and only has a paramterized default member.
            AssertInputCodeYieldsExpectedInspectionResultCount(input, expectResultCount, ReferenceLibrary.AdoDb);
        }

        [Test]
        [Category("Inspections")]
        public void ObjectVariableNotSet_NewExprWithNonObjectDefaultMember_NoResult()
        {
            var expectResultCount = 0;
            var input =
                @"
Private Sub Test()    
    Dim foo As Variant  
    foo = New Connection
End Sub";
            AssertInputCodeYieldsExpectedInspectionResultCount(input, expectResultCount, ReferenceLibrary.AdoDb);
        }

        [Test]
        [Category("Inspections")]
        public void ObjectVariableNotSet_NewExprWithObjectOnlyDefaultMember_OneResult()
        {
            var expectResultCount = 1;
            var input =
                @"
Private Sub Test()    
    Dim foo As Variant  
    foo = New Recordset
End Sub";
            //The default member of Recordset is Fields, which is an object.
            AssertInputCodeYieldsExpectedInspectionResultCount(input, expectResultCount, ReferenceLibrary.AdoDb);
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

        [Test]
        [Category("Grammar")]
        [Category("Resolver")]
        [TestCase("    Bar cls.Baz")]
        [TestCase("    Baz (cls.Baz)")]
        [TestCase("    Debug.Print cls.Baz")]
        [TestCase("    Debug.Print 42, cls.Baz")]
        [TestCase("    Debug.Print 42; cls.Baz")]
        [TestCase("    Debug.Print Spc(cls.Baz)")]
        [TestCase("    Debug.Print 42, Spc(cls.Baz)")]
        [TestCase("    Debug.Print 42; Spc(cls.Baz)")]
        [TestCase("    Debug.Print Tab(cls.Baz)")]
        [TestCase("    Debug.Print 42, Tab(cls.Baz)")]
        [TestCase("    Debug.Print 42; Tab(cls.Baz)")]
        [TestCase("    If cls.Baz Then Foo = 42")]
        [TestCase("    If cls.Baz Then \r\n        Foo = 42 \r\n    End If")]
        [TestCase("    If False Then : ElseIf cls.Baz Then\r\n        Foo = 42 \r\n    End If")]
        [TestCase("    Do While cls.Baz\r\n        Foo = 42 \r\n    Loop")]
        [TestCase("    Do Until cls.Baz\r\n        Foo = 42 \r\n    Loop")]
        [TestCase("    Do : Foo = 42 :  Loop While cls.Baz")]
        [TestCase("    Do : Foo = 42 :  Loop Until cls.Baz")]
        [TestCase("    While cls.Baz\r\n        Foo = 42 \r\n    Wend")]
        [TestCase("    For fooBar = cls.Baz To 42 Step 23\r\n        Foo = 42 \r\n    Next")]
        [TestCase("    For fooBar = 42 To cls.Baz Step 23\r\n        Foo = 42 \r\n    Next")]
        [TestCase("    For fooBar = 23 To 42 Step cls.Baz\r\n        Foo = 42 \r\n    Next")]
        [TestCase("    Select Case cls.Baz : Case 42 : Foo = 42 : End Select")]
        [TestCase("    Select Case 42 : Case cls.Baz : Foo = 42 : End Select")]
        [TestCase("    Select Case 42 : Case 23, cls.Baz : Foo = 42 : End Select")]
        [TestCase("    Select Case 42 : Case cls.Baz To 666 : Foo = 42 : End Select")]
        [TestCase("    Select Case 42 : Case 23 To cls.Baz : Foo = 42 : End Select")]
        [TestCase("    Select Case 42 : Case Is = cls.Baz : Foo = 42 : End Select")]
        [TestCase("    Select Case 42 : Case Is < cls.Baz : Foo = 42 : End Select")]
        [TestCase("    Select Case 42 : Case Is > cls.Baz : Foo = 42 : End Select")]
        [TestCase("    Select Case 42 : Case Is <> cls.Baz : Foo = 42 : End Select")]
        [TestCase("    Select Case 42 : Case Is >< cls.Baz : Foo = 42 : End Select")]
        [TestCase("    Select Case 42 : Case Is <= cls.Baz : Foo = 42 : End Select")]
        [TestCase("    Select Case 42 : Case Is =< cls.Baz : Foo = 42 : End Select")]
        [TestCase("    Select Case 42 : Case Is >= cls.Baz : Foo = 42 : End Select")]
        [TestCase("    Select Case 42 : Case Is => cls.Baz : Foo = 42 : End Select")]
        [TestCase("    Select Case 42 : Case = cls.Baz : Foo = 42 : End Select")]
        [TestCase("    Select Case 42 : Case < cls.Baz : Foo = 42 : End Select")]
        [TestCase("    Select Case 42 : Case > cls.Baz : Foo = 42 : End Select")]
        [TestCase("    Select Case 42 : Case <> cls.Baz : Foo = 42 : End Select")]
        [TestCase("    Select Case 42 : Case >< cls.Baz : Foo = 42 : End Select")]
        [TestCase("    Select Case 42 : Case <= cls.Baz : Foo = 42 : End Select")]
        [TestCase("    Select Case 42 : Case =< cls.Baz : Foo = 42 : End Select")]
        [TestCase("    Select Case 42 : Case >= cls.Baz : Foo = 42 : End Select")]
        [TestCase("    Select Case 42 : Case => cls.Baz : Foo = 42 : End Select")]
        [TestCase("    On cls.Baz GoTo label1, label2")]
        [TestCase("    On cls.Baz GoSub label1, label2")]
        [TestCase("    ReDim fooBar(cls.Baz To 42)")]
        [TestCase("    ReDim fooBar(23 To cls.Baz)")]
        [TestCase("    ReDim fooBar(23 To 42, cls.Baz To 42)")]
        [TestCase("    ReDim fooBar(23 To 42, 23 To cls.Baz)")]
        [TestCase("    ReDim fooBar(cls.Baz)")]
        [TestCase("    ReDim fooBar(42, cls.Baz)")]
        [TestCase("    Mid(fooBar, cls.Baz, 42) = \"Hello\"")]
        [TestCase("    Mid(fooBar, 23, cls.Baz) = \"Hello\"")]
        [TestCase("    Mid(fooBar, 23, 42) = cls.Baz")]
        [TestCase("    LSet fooBar = cls.Baz")]
        [TestCase("    RSet fooBar = cls.Baz")]
        [TestCase("    Error cls.Baz")]
        [TestCase("    Open cls.Baz As 42 Len = 23")]
        [TestCase("    Open \"somePath\" As cls.Baz Len = 23")]
        [TestCase("    Open \"somePath\" As #cls.Baz Len = 23")]
        [TestCase("    Open \"somePath\" As 23 Len = cls.Baz")]
        [TestCase("    Close cls.Baz, 23")]
        [TestCase("    Close 23, #cls.Baz, 23")]
        [TestCase("    Seek cls.Baz, 23")]
        [TestCase("    Seek #cls.Baz, 23")]
        [TestCase("    Seek 23, cls.Baz")]
        [TestCase("    Lock cls.Baz, 23 To 42")]
        [TestCase("    Lock #cls.Baz, 23 To 42")]
        [TestCase("    Lock 23, cls.Baz To 42")]
        [TestCase("    Lock 23, 42 To cls.Baz")]
        [TestCase("    Unlock cls.Baz, 23 To 42")]
        [TestCase("    Unlock #cls.Baz, 23 To 42")]
        [TestCase("    Unlock 23, cls.Baz To 42")]
        [TestCase("    Unlock 23, 42 To cls.Baz")]
        [TestCase("    Line Input #cls.Baz, fooBar")]
        [TestCase("    Width #cls.Baz, 42")]
        [TestCase("    Width #23, cls.Baz")]
        [TestCase("    Print #cls.Baz, 42")]
        [TestCase("    Print #23, cls.Baz")]
        [TestCase("    Print #23, 42, cls.Baz")]
        [TestCase("    Print #23, 42; cls.Baz")]
        [TestCase("    Print #23, Spc(cls.Baz)")]
        [TestCase("    Print #23, 42, Spc(cls.Baz)")]
        [TestCase("    Print #23, 42; Spc(cls.Baz)")]
        [TestCase("    Print #23, Tab(cls.Baz)")]
        [TestCase("    Print #23, 42, Tab(cls.Baz)")]
        [TestCase("    Print #23, 42; Tab(cls.Baz)")]
        [TestCase("    Input #cls.Baz, fooBar")]
        [TestCase("    Put cls.Baz, 42, fooBar")]
        [TestCase("    Put #cls.Baz, 42, fooBar")]
        [TestCase("    Put 42, cls.Baz, fooBar")]
        [TestCase("    Get cls.Baz, 42, fooBar")]
        [TestCase("    Get #cls.Baz, 42, fooBar")]
        [TestCase("    Get 42, cls.Baz, fooBar")]
        [TestCase("    Name \"somePath\" As cls.Baz")]
        [TestCase("    Name cls.Baz As \"somePath\"")]
        public void FailedLetCoercionNotInLetStatement_NoResult(string statement)
        {
            var class1Code = @"
Public Function Foo() As Long
End Function
";

            var class2Code = @"
Public Function Baz() As Class1
Attribute Baz.VB_UserMemId = 0
    Set Baz = New Class1
End Function
";

            var moduleCode = $@"
Private Function Foo() As Long 
    Dim cls As new Class2
    Dim fooBar As Variant
{statement}
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
        [Category("Grammar")]
        [Category("Resolver")]
        [TestCase("    Foo = cls.Baz")]
        [TestCase("    Let Foo = cls.Baz")]
        [TestCase("    Foo = cls.Baz + 42")]
        [TestCase("    Foo = cls.Baz - 42")]
        [TestCase("    Foo = cls.Baz * 42")]
        [TestCase("    Foo = cls.Baz ^ 42")]
        [TestCase("    Foo = cls.Baz \\ 42")]
        [TestCase("    Foo = cls.Baz Mod 42")]
        [TestCase("    Foo = cls.Baz & \" sheep\"")]
        [TestCase("    Foo = cls.Baz And 42")]
        [TestCase("    Foo = cls.Baz Or 42")]
        [TestCase("    Foo = cls.Baz Xor 42")]
        [TestCase("    Foo = cls.Baz Eqv 42")]
        [TestCase("    Foo = cls.Baz Imp 42")]
        [TestCase("    Foo = cls.Baz = 42")]
        [TestCase("    Foo = cls.Baz < 42")]
        [TestCase("    Foo = cls.Baz > 42")]
        [TestCase("    Foo = cls.Baz <= 42")]
        [TestCase("    Foo = cls.Baz =< 42")]
        [TestCase("    Foo = cls.Baz >= 42")]
        [TestCase("    Foo = cls.Baz => 42")]
        [TestCase("    Foo = cls.Baz <> 42")]
        [TestCase("    Foo = cls.Baz >< 42")]
        [TestCase("    Foo = cls.Baz Like \"Hello\"")]
        [TestCase("    Foo = 42 + cls.Baz")]
        [TestCase("    Foo = 42 * cls.Baz")]
        [TestCase("    Foo = 42 - cls.Baz")]
        [TestCase("    Foo = 42 ^ cls.Baz")]
        [TestCase("    Foo = 42 \\ cls.Baz")]
        [TestCase("    Foo = 42 Mod cls.Baz")]
        [TestCase("    Foo = \"sheep\" & cls.Baz")]
        [TestCase("    Foo = 42 And cls.Baz")]
        [TestCase("    Foo = 42 Or cls.Baz")]
        [TestCase("    Foo = 42 Xor cls.Baz")]
        [TestCase("    Foo = 42 Eqv cls.Baz")]
        [TestCase("    Foo = 42 Imp cls.Baz")]
        [TestCase("    Foo = 42 = cls.Baz")]
        [TestCase("    Foo = 42 < cls.Baz")]
        [TestCase("    Foo = 42 > cls.Baz")]
        [TestCase("    Foo = 42 <= cls.Baz")]
        [TestCase("    Foo = 42 =< cls.Baz")]
        [TestCase("    Foo = 42 >= cls.Baz")]
        [TestCase("    Foo = 42 => cls.Baz")]
        [TestCase("    Foo = 42 <> cls.Baz")]
        [TestCase("    Foo = 42 >< cls.Baz")]
        [TestCase("    Foo = \"Hello\" Like cls.Baz")]
        [TestCase("    Foo = -cls.Baz")]
        [TestCase("    Foo = Not cls.Baz")]
        public void FailedLetCoercionInLetStatementWithLHSValueType_NoResult(string statement)
        {
            var class1Code = @"
Public Function Foo() As Long
End Function
";

            var class2Code = @"
Public Function Baz() As Class1
Attribute Baz.VB_UserMemId = 0
    Set Baz = New Class1
End Function
";

            var moduleCode = $@"
Private Function Foo() As Long 
    Dim cls As new Class2
    Dim fooBar As Variant
{statement}
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
        [Category("Grammar")]
        [Category("Resolver")]
        [TestCase("    Foo = cls.Baz + 42")]
        [TestCase("    Foo = cls.Baz - 42")]
        [TestCase("    Foo = cls.Baz * 42")]
        [TestCase("    Foo = cls.Baz ^ 42")]
        [TestCase("    Foo = cls.Baz \\ 42")]
        [TestCase("    Foo = cls.Baz Mod 42")]
        [TestCase("    Foo = cls.Baz & \" sheep\"")]
        [TestCase("    Foo = cls.Baz And 42")]
        [TestCase("    Foo = cls.Baz Or 42")]
        [TestCase("    Foo = cls.Baz Xor 42")]
        [TestCase("    Foo = cls.Baz Eqv 42")]
        [TestCase("    Foo = cls.Baz Imp 42")]
        [TestCase("    Foo = cls.Baz = 42")]
        [TestCase("    Foo = cls.Baz < 42")]
        [TestCase("    Foo = cls.Baz > 42")]
        [TestCase("    Foo = cls.Baz <= 42")]
        [TestCase("    Foo = cls.Baz =< 42")]
        [TestCase("    Foo = cls.Baz >= 42")]
        [TestCase("    Foo = cls.Baz => 42")]
        [TestCase("    Foo = cls.Baz <> 42")]
        [TestCase("    Foo = cls.Baz >< 42")]
        [TestCase("    Foo = cls.Baz Like \"Hello\"")]
        [TestCase("    Foo = 42 + cls.Baz")]
        [TestCase("    Foo = 42 * cls.Baz")]
        [TestCase("    Foo = 42 - cls.Baz")]
        [TestCase("    Foo = 42 ^ cls.Baz")]
        [TestCase("    Foo = 42 \\ cls.Baz")]
        [TestCase("    Foo = 42 Mod cls.Baz")]
        [TestCase("    Foo = \"sheep\" & cls.Baz")]
        [TestCase("    Foo = 42 And cls.Baz")]
        [TestCase("    Foo = 42 Or cls.Baz")]
        [TestCase("    Foo = 42 Xor cls.Baz")]
        [TestCase("    Foo = 42 Eqv cls.Baz")]
        [TestCase("    Foo = 42 Imp cls.Baz")]
        [TestCase("    Foo = 42 = cls.Baz")]
        [TestCase("    Foo = 42 < cls.Baz")]
        [TestCase("    Foo = 42 > cls.Baz")]
        [TestCase("    Foo = 42 <= cls.Baz")]
        [TestCase("    Foo = 42 =< cls.Baz")]
        [TestCase("    Foo = 42 >= cls.Baz")]
        [TestCase("    Foo = 42 => cls.Baz")]
        [TestCase("    Foo = 42 <> cls.Baz")]
        [TestCase("    Foo = 42 >< cls.Baz")]
        [TestCase("    Foo = \"Hello\" Like cls.Baz")]
        [TestCase("    Foo = -cls.Baz")]
        [TestCase("    Foo = Not cls.Baz")]
        public void FailedLetCoercionInLetStatementButNotOnEntireRHSWithLHSDefaultMemberCall_NoResult(string statement)
        {
            var class1Code = @"
Public Function Foo() As Long
End Function
";

            var class2Code = @"
Public Function Baz() As Class1
Attribute Baz.VB_UserMemId = 0
    Set Baz = New Class1
End Function
";
            var class3Code = @"
Public Function Foo() As Long
Attribute Foo.VB_UserMemId = 0
End Function
";

            var moduleCode = $@"
Private Function FooBar() As Variant  
    Dim cls As new Class2
    Dim Foo As Class3
    Dim fooBaz() As Class3
{statement}
End Function

Private Sub Bar(arg As Long)
End Sub

Private Sub Baz(arg As Variant)
End Sub
";

            var inspectionResults = InspectionResultsForModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Class2", class2Code, ComponentType.ClassModule),
                ("Class3", class3Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Grammar")]
        [Category("Resolver")]
        [TestCase("    Foo = cls.Baz + 42")]
        [TestCase("    Foo = cls.Baz - 42")]
        [TestCase("    Foo = cls.Baz * 42")]
        [TestCase("    Foo = cls.Baz ^ 42")]
        [TestCase("    Foo = cls.Baz \\ 42")]
        [TestCase("    Foo = cls.Baz Mod 42")]
        [TestCase("    Foo = cls.Baz & \" sheep\"")]
        [TestCase("    Foo = cls.Baz And 42")]
        [TestCase("    Foo = cls.Baz Or 42")]
        [TestCase("    Foo = cls.Baz Xor 42")]
        [TestCase("    Foo = cls.Baz Eqv 42")]
        [TestCase("    Foo = cls.Baz Imp 42")]
        [TestCase("    Foo = cls.Baz = 42")]
        [TestCase("    Foo = cls.Baz < 42")]
        [TestCase("    Foo = cls.Baz > 42")]
        [TestCase("    Foo = cls.Baz <= 42")]
        [TestCase("    Foo = cls.Baz =< 42")]
        [TestCase("    Foo = cls.Baz >= 42")]
        [TestCase("    Foo = cls.Baz => 42")]
        [TestCase("    Foo = cls.Baz <> 42")]
        [TestCase("    Foo = cls.Baz >< 42")]
        [TestCase("    Foo = cls.Baz Like \"Hello\"")]
        [TestCase("    Foo = 42 + cls.Baz")]
        [TestCase("    Foo = 42 * cls.Baz")]
        [TestCase("    Foo = 42 - cls.Baz")]
        [TestCase("    Foo = 42 ^ cls.Baz")]
        [TestCase("    Foo = 42 \\ cls.Baz")]
        [TestCase("    Foo = 42 Mod cls.Baz")]
        [TestCase("    Foo = \"sheep\" & cls.Baz")]
        [TestCase("    Foo = 42 And cls.Baz")]
        [TestCase("    Foo = 42 Or cls.Baz")]
        [TestCase("    Foo = 42 Xor cls.Baz")]
        [TestCase("    Foo = 42 Eqv cls.Baz")]
        [TestCase("    Foo = 42 Imp cls.Baz")]
        [TestCase("    Foo = 42 = cls.Baz")]
        [TestCase("    Foo = 42 < cls.Baz")]
        [TestCase("    Foo = 42 > cls.Baz")]
        [TestCase("    Foo = 42 <= cls.Baz")]
        [TestCase("    Foo = 42 =< cls.Baz")]
        [TestCase("    Foo = 42 >= cls.Baz")]
        [TestCase("    Foo = 42 => cls.Baz")]
        [TestCase("    Foo = 42 <> cls.Baz")]
        [TestCase("    Foo = 42 >< cls.Baz")]
        [TestCase("    Foo = \"Hello\" Like cls.Baz")]
        [TestCase("    Foo = -cls.Baz")]
        [TestCase("    Foo = Not cls.Baz")]
        public void FailedLetCoercionInLetStatementButNotOnEntireRHSWithLHSUnboundDefaultMemberCall_NoResult(string statement)
        {
            var class1Code = @"
Public Function Foo() As Long
End Function
";

            var class2Code = @"
Public Function Baz() As Class1
Attribute Baz.VB_UserMemId = 0
    Set Baz = New Class1
End Function
";
            var class3Code = @"
Public Function Foo() As Long
Attribute Foo.VB_UserMemId = 0
End Function
";

            var moduleCode = $@"
Private Function FooBar() As Variant  
    Dim cls As new Class2
    Dim Foo As Object
    Dim fooBaz() As Object
{statement}
End Function

Private Sub Bar(arg As Long)
End Sub

Private Sub Baz(arg As Variant)
End Sub
";

            var inspectionResults = InspectionResultsForModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Class2", class2Code, ComponentType.ClassModule),
                ("Class3", class3Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Grammar")]
        [Category("Resolver")]
        [TestCase("    Foo = cls.Baz + 42")]
        [TestCase("    Foo = cls.Baz - 42")]
        [TestCase("    Foo = cls.Baz * 42")]
        [TestCase("    Foo = cls.Baz ^ 42")]
        [TestCase("    Foo = cls.Baz \\ 42")]
        [TestCase("    Foo = cls.Baz Mod 42")]
        [TestCase("    Foo = cls.Baz & \" sheep\"")]
        [TestCase("    Foo = cls.Baz And 42")]
        [TestCase("    Foo = cls.Baz Or 42")]
        [TestCase("    Foo = cls.Baz Xor 42")]
        [TestCase("    Foo = cls.Baz Eqv 42")]
        [TestCase("    Foo = cls.Baz Imp 42")]
        [TestCase("    Foo = cls.Baz = 42")]
        [TestCase("    Foo = cls.Baz < 42")]
        [TestCase("    Foo = cls.Baz > 42")]
        [TestCase("    Foo = cls.Baz <= 42")]
        [TestCase("    Foo = cls.Baz =< 42")]
        [TestCase("    Foo = cls.Baz >= 42")]
        [TestCase("    Foo = cls.Baz => 42")]
        [TestCase("    Foo = cls.Baz <> 42")]
        [TestCase("    Foo = cls.Baz >< 42")]
        [TestCase("    Foo = cls.Baz Like \"Hello\"")]
        [TestCase("    Foo = 42 + cls.Baz")]
        [TestCase("    Foo = 42 * cls.Baz")]
        [TestCase("    Foo = 42 - cls.Baz")]
        [TestCase("    Foo = 42 ^ cls.Baz")]
        [TestCase("    Foo = 42 \\ cls.Baz")]
        [TestCase("    Foo = 42 Mod cls.Baz")]
        [TestCase("    Foo = \"sheep\" & cls.Baz")]
        [TestCase("    Foo = 42 And cls.Baz")]
        [TestCase("    Foo = 42 Or cls.Baz")]
        [TestCase("    Foo = 42 Xor cls.Baz")]
        [TestCase("    Foo = 42 Eqv cls.Baz")]
        [TestCase("    Foo = 42 Imp cls.Baz")]
        [TestCase("    Foo = 42 = cls.Baz")]
        [TestCase("    Foo = 42 < cls.Baz")]
        [TestCase("    Foo = 42 > cls.Baz")]
        [TestCase("    Foo = 42 <= cls.Baz")]
        [TestCase("    Foo = 42 =< cls.Baz")]
        [TestCase("    Foo = 42 >= cls.Baz")]
        [TestCase("    Foo = 42 => cls.Baz")]
        [TestCase("    Foo = 42 <> cls.Baz")]
        [TestCase("    Foo = 42 >< cls.Baz")]
        [TestCase("    Foo = \"Hello\" Like cls.Baz")]
        [TestCase("    Foo = -cls.Baz")]
        [TestCase("    Foo = Not cls.Baz")]
        public void FailedLetCoercionInLetStatementButNotOnEntireRHSWithLHSVariantCall_NoResult(string statement)
        {
            var class1Code = @"
Public Function Foo() As Long
End Function
";

            var class2Code = @"
Public Function Baz() As Class1
Attribute Baz.VB_UserMemId = 0
    Set Baz = New Class1
End Function
";
            var class3Code = @"
Public Function Foo() As Long
Attribute Foo.VB_UserMemId = 0
End Function
";

            var moduleCode = $@"
Private Function FooBar() As Variant  
    Dim cls As new Class2
    Dim Foo As Variant
    Dim fooBaz() As Variant
{statement}
End Function

Private Sub Bar(arg As Long)
End Sub

Private Sub Baz(arg As Variant)
End Sub
";

            var inspectionResults = InspectionResultsForModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Class2", class2Code, ComponentType.ClassModule),
                ("Class3", class3Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Grammar")]
        [Category("Resolver")]
        [TestCase("    Foo = cls.Baz", 5, 8)]
        [TestCase("    Let Foo = cls.Baz", 9, 12)]
        [TestCase("    fooBaz(42) = cls.Baz", 5, 15)]
        [TestCase("    Let fooBaz(42) = cls.Baz", 9, 19)]
        [TestCase("    Foo = cls.Bar(23)", 5, 8)]
        [TestCase("    Let Foo = cls.Bar(23)", 9, 12)]
        [TestCase("    fooBaz(42) = cls.Bar(23)", 5, 15)]
        [TestCase("    Let fooBaz(42) = cls.Bar(23)", 9, 19)]
        public void FailedLetCoercionInLetStatementOnEntireRHSWithLHSDefaultMemberCall_OneResult(string statement, int selectionStartColumn, int selectionEndColumn)
        {
            var class1Code = @"
Public Function Foo() As Long
End Function
";

            var class2Code = @"
Public Function Baz() As Class1
Attribute Baz.VB_UserMemId = 0
    Set Baz = New Class1
End Function

Public Function Bar() As Class1()
End Function
";
            var class3Code = @"
Public Function Foo() As Long
Attribute Baz.VB_UserMemId = 0
End Function
";

            var moduleCode = $@"
Private Function FooBar() As Variant  
    Dim cls As new Class2
    Dim Foo As Class3
    Dim fooBaz() As Class3
{statement}
End Function

Private Sub Bar(arg As Long)
End Sub

Private Sub Baz(arg As Variant)
End Sub
";

            var inspectionResults = InspectionResultsForModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Class2", class2Code, ComponentType.ClassModule),
                ("Class3", class3Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            var inspectionResult = inspectionResults.Single();

            var expectedSelection = new Selection(6, selectionStartColumn, 6, selectionEndColumn);
            var actualSelection = inspectionResult.QualifiedSelection.Selection;

            Assert.AreEqual(expectedSelection, actualSelection);
        }

        [Test]
        [Category("Grammar")]
        [Category("Resolver")]
        [TestCase("    Foo = cls.Baz", 5, 8)]
        [TestCase("    Let Foo = cls.Baz", 9, 12)]
        [TestCase("    fooBaz(42) = cls.Baz", 5, 15)]
        [TestCase("    Let fooBaz(42) = cls.Baz", 9, 19)]
        [TestCase("    Foo = cls.Bar(23)", 5, 8)]
        [TestCase("    Let Foo = cls.Bar(23)", 9, 12)]
        [TestCase("    fooBaz(42) = cls.Bar(23)", 5, 15)]
        [TestCase("    Let fooBaz(42) = cls.Bar(23)", 9, 19)]
        public void FailedLetCoercionInLetStatementOnEntireRHSWithLHSUnboundDefaultMemberCall_OneResult(string statement, int selectionStartColumn, int selectionEndColumn)
        {
            var class1Code = @"
Public Function Foo() As Long
End Function
";

            var class2Code = @"
Public Function Baz() As Class1
Attribute Baz.VB_UserMemId = 0
    Set Baz = New Class1
End Function

Public Function Bar() As Class1()
End Function
";
            var class3Code = @"
Public Function Foo() As Long
Attribute Baz.VB_UserMemId = 0
End Function
";

            var moduleCode = $@"
Private Function FooBar() As Variant  
    Dim cls As new Class2
    Dim Foo As Object
    Dim fooBaz() As Object
{statement}
End Function

Private Sub Bar(arg As Long)
End Sub

Private Sub Baz(arg As Variant)
End Sub
";

            var inspectionResults = InspectionResultsForModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Class2", class2Code, ComponentType.ClassModule),
                ("Class3", class3Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            var inspectionResult = inspectionResults.Single();

            var expectedSelection = new Selection(6, selectionStartColumn, 6, selectionEndColumn);
            var actualSelection = inspectionResult.QualifiedSelection.Selection;

            Assert.AreEqual(expectedSelection, actualSelection);
        }

        [Test]
        [Category("Grammar")]
        [Category("Resolver")]
        [TestCase("    Foo = cls.Baz", 5, 8)]
        [TestCase("    Let Foo = cls.Baz", 9, 12)]
        [TestCase("    fooBaz(42) = cls.Baz", 5, 15)]
        [TestCase("    Let fooBaz(42) = cls.Baz", 9, 19)]
        [TestCase("    Foo = cls.Bar(23)", 5, 8)]
        [TestCase("    Let Foo = cls.Bar(23)", 9, 12)]
        [TestCase("    fooBaz(42) = cls.Bar(23)", 5, 15)]
        [TestCase("    Let fooBaz(42) = cls.Bar(23)", 9, 19)]
        public void FailedLetCoercionInLetStatementOnEntireRHSWithLHSVariantCall_OneResult(string statement, int selectionStartColumn, int selectionEndColumn)
        {
            var class1Code = @"
Public Function Foo() As Long
End Function
";

            var class2Code = @"
Public Function Baz() As Class1
Attribute Baz.VB_UserMemId = 0
    Set Baz = New Class1
End Function

Public Function Bar() As Class1()
End Function
";
            var class3Code = @"
Public Function Foo() As Long
Attribute Baz.VB_UserMemId = 0
End Function
";

            var moduleCode = $@"
Private Function FooBar() As Variant  
    Dim cls As new Class2
    Dim Foo As Variant
    Dim fooBaz() As Variant
{statement}
End Function

Private Sub Bar(arg As Long)
End Sub

Private Sub Baz(arg As Variant)
End Sub
";

            var inspectionResults = InspectionResultsForModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Class2", class2Code, ComponentType.ClassModule),
                ("Class3", class3Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            var inspectionResult = inspectionResults.Single();

            var expectedSelection = new Selection(6, selectionStartColumn, 6, selectionEndColumn);
            var actualSelection = inspectionResult.QualifiedSelection.Selection;

            Assert.AreEqual(expectedSelection, actualSelection);
        }

        [Test]
        [Category("Grammar")]
        [Category("Resolver")]
        [TestCase("    cls.Baz = 42", 5, 12)]
        [TestCase("    Let cls.Baz = 42", 9, 16)]
        [TestCase("    cls.Bar(42) = 42", 5, 16)]
        [TestCase("    Let cls.Bar(42) = 42", 9, 20)]
        public void FailedLetCoercionAssignmentOnLHSOfLetStatement_OneResult(string statement, int selectionStartColumn, int selectionEndColumn)
        {
            var class1Code = @"
Public Property Let Foo(arg As Long)
End Property
";

            var class2Code = @"
Public Function Baz() As Class1
Attribute Baz.VB_UserMemId = 0
    Set Baz = New Class1
End Function

Public Function Bar() As Class1()
End Function
";

            var moduleCode = $@"
Private Function Foo() As Variant 
    Dim cls As new Class2
{statement}
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

            var inspectionResult = inspectionResults.Single();

            var expectedSelection = new Selection(4, selectionStartColumn, 4, selectionEndColumn);
            var actualSelection = inspectionResult.QualifiedSelection.Selection;

            Assert.AreEqual(expectedSelection, actualSelection);
        }

        [Test]
        [Category("Grammar")]
        [Category("Resolver")]
        [TestCase("    cls.Baz = fooBar")]
        [TestCase("    Let cls.Baz = fooBar")]
        //This prevents problems with some types in libraries like OLE_COLOR, which are not really classes.
        //See issue #4997 at https://github.com/rubberduck-vba/Rubberduck/issues/4997
        public void PropertyLetOnLHS_NoResult(string statement)
        {
            var class1Code = @"
Public Function Foo() As Long
Attribute Foo.VB_UserMemId = 0
End Function
";

            var class2Code = @"
Public Property Let Baz(RHS As Class1)
End Property

Public Property Get Baz() As Class1
Attribute Baz.VB_UserMemId = 0
End Property
";

            var moduleCode = $@"
Private Function Foo() As Variant 
    Dim cls As New Class2
    Dim fooBar As New Class1
{statement}
End Function
";

            var inspectionResults = InspectionResultsForModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Class2", class2Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            Assert.IsFalse(inspectionResults.Any());
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new ObjectVariableNotSetInspection(state);
        }

        private void AssertInputCodeYieldsExpectedInspectionResultCount(string inputCode, int expected, params ReferenceLibrary[] testLibraries)
        {
            var inspectionResults = InspectionResultsForModules(("Class1", inputCode, ComponentType.ClassModule), testLibraries);
            Assert.AreEqual(expected, inspectionResults.Count());
        }
    }
}

