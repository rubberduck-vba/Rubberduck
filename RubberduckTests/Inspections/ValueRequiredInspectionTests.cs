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
    public class ValueRequiredInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
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
        public void FailedLetCoercionNotInLetStatement_OneResult(string statement)
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

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
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
        public void FailedLetCoercionInLetStatementInRHS_OneResult(string statement)
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

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        [TestCase("    Foo = cls.Baz")]
        [TestCase("    Let Foo = cls.Baz")]
        [TestCase("    fooBaz(42) = cls.Baz")]
        [TestCase("    Let fooBaz(42) = cls.Baz")]
        [TestCase("    Foo = cls.Bar(23)")]
        [TestCase("    Let Foo = cls.Bar(23)")]
        [TestCase("    fooBaz(42) = cls.Bar(23)")]
        [TestCase("    Let fooBaz(42) = cls.Bar(23)")]
        public void FailedLetCoercionInLetStatementOnEntireRHS_OneResult(string statement)
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

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        [TestCase("    cls.Baz = 42")]
        [TestCase("    Let cls.Baz = 42")]
        [TestCase("    cls.Bar(42) = 42")]
        [TestCase("    Let cls.Bar(42) = 42")]
        public void FailedLetCoercionAssignment_NoResult(string statement)
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

            Assert.IsFalse(inspectionResults.Any());
        }

        [Category("Inspections")]
        [Test]
        public void ParamArray_NoResult()
        {
            var classCode = @"
Public Function Foo(index As Variant) As Class1
End Function
";

            var moduleCode = $@"
Private Function Foo(ParamArray args() As Variant) As Variant
End Function

Private Function Test() As Variant
    Dim bar As Class1
    Set bar = New Class1
    Test = Foo(bar)
End Function
";

            var inspectionResults = InspectionResultsForModules(
                ("Class1", classCode, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            Assert.IsFalse(inspectionResults.Any());
        }

        [Category("Inspections")]
        [Test]
        public void ParamArrayInLibrary_NoResult()
        {
            var classCode = @"
Public Function Foo(index As Variant) As Class1
End Function
";

            var moduleCode = $@"
Private Function Test() As Variant
    Dim bar As Class1
    Set bar = New Class1
    Test = Array(bar)
End Function
";

            var modules = new (string, string, ComponentType)[] 
            {
                ("Class1", classCode, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule),
            };

            Assert.IsFalse(InspectionResultsForModules(modules, ReferenceLibrary.VBA).Any());
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new ValueRequiredInspection(state);
        }
    }
}