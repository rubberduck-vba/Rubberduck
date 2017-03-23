using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RubberduckTests.Inspections
{
    public static class ObjectVariableNotSetInspectionTests_VBABlocks
    {
        public static KeyValuePair<string, int> GivenIndexerObjectAccess_ReturnsNoResult_TestParams()
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
            return new KeyValuePair<string, int>(input, expectResultCount);
        }

        public static KeyValuePair<string, int> GivenIndexerObjectAccess_ReturnsResult_TestParams()
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
            return new KeyValuePair<string, int>(input, expectResultCount);
        }

        public static KeyValuePair<string, int> GivenStringVariable_ReturnsNoResult_TestParams()
        {
            var expectResultCount = 0;
            var input =
@"
Private Sub Workbook_Open()
    
    Dim target As String
    target = Range(""A1"")
    
    target.Value = ""all good""

End Sub";
            return new KeyValuePair<string, int>(input, expectResultCount);
        }

        public static KeyValuePair<string, int> GivenVariantVariableAssignedObject_ReturnsResult_TestParams()
        {
            var expectResultCount = 1;
            var input =
@"
Private Sub TestSub(ByRef testParam As Variant)
'whoCares is a LExprContext and is a known interesting declaration
    Dim target As Collection
    Set target = new Collection
    testParam = target             
    testParam.Add 100
End Sub";
            return new KeyValuePair<string, int>(input, expectResultCount);
        }

        public static KeyValuePair<string, int> GivenVariantVariableAssignedNewObject_ReturnsResult_TestParams()
        {
            var expectResultCount = 1;
            var input =
@"
Private Sub TestSub(ByRef testParam As Variant)
'is a NewExprContext
    testParam = new Collection     
End Sub";
            return new KeyValuePair<string, int>(input, expectResultCount);
        }

        public static KeyValuePair<string, int> GivenVariantVariableAssignedRange_ReturnsResult_TestParams()
        {
            var expectResultCount = 1;
            var input =
@"
Private Sub TestSub(ByRef testParam As Variant)
'Range(""A1:C1"") is a LExprContext but is not a known interesting declaration
    testParam = Range(""A1:C1"")    
End Sub";
            return new KeyValuePair<string, int>(input, expectResultCount);
        }

        public static KeyValuePair<string, int> GivenVariantVariableAssignedDeclaredRange_ReturnsResult_TestParams()
        {
            var expectResultCount = 1;
            var input =
@"
Private Sub TestSub(ByRef testParam As Variant, target As Range)
'target is a LExprContext and is a known interesting declaration
    testParam = target
End Sub";
            return new KeyValuePair<string, int>(input, expectResultCount);
        }

        public static KeyValuePair<string, int> GivenVariantVariableAssignedDeclaredVariant_ReturnsNoResult_TestParams()
        {
            var expectResultCount = 0;
            var input =
@"
Private Sub TestSub(ByRef testParam As Variant, target As Variant)
'target is a LExprContext, is a known interesting declaration - but is a Variant
    testParam = target           
End Sub";
            return new KeyValuePair<string, int>(input, expectResultCount);
        }

        public static KeyValuePair<string, int> GivenVariantVariableAssignedBaseType_ReturnsNoResult_TestParams()
        {
            var expectResultCount = 0;
            var input =
@"
Private Sub Workbook_Open()
    Dim target As Variant
    target = ""A1""     'is a LiteralExprContext
End Sub";
            return new KeyValuePair<string, int>(input, expectResultCount);
        }

        public static KeyValuePair<string, int> GivenObjectVariableNotSet_ReturnsResult_TestParams()
        {
            var expectResultCount = 1;
            var input =
@"
Private Sub Workbook_Open()
    
    Dim target As Range
    target = Range(""A1"")
    
    target.Value = ""forgot something?""

End Sub";
            return new KeyValuePair<string, int>(input, expectResultCount);
        }

        public static KeyValuePair<string, int> GivenObjectVariableNotSet_Ignored_DoesNotReturnResult_TestParams()
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
            return new KeyValuePair<string, int>(input, expectResultCount);
        }

        public static KeyValuePair<string, int> GivenSetObjectVariable_ReturnsNoResult_TestParams()
        {
            var expectResultCount = 0;
            var input =
@"
Private Sub Workbook_Open()
    
    Dim target As Range
    Set target = Range(""A1"")
    
    target.Value = ""All good""

End Sub";
            return new KeyValuePair<string, int>(input, expectResultCount);
        }

        public static KeyValuePair<string, int> FunctionReturnsArrayOfType_ReturnsNoResult_TestParams()
        {
            var expectedResultCount = 0;
            var input =
@"
Private Function GetSomeDictionaries() As Dictionary()
    Dim temp(0 To 1) As Worksheet
    Set temp(0) = New Dictionary
    Set temp(1) = New Dictionary
    GetSomeDictionaries = temp
End Function";
            return new KeyValuePair<string, int>(input, expectedResultCount);
        }

        public static string IgnoreQuickFixWorks_Input()
        {
            return
@"
Private Sub Workbook_Open()
    
    Dim target As Range
    target = Range(""A1"")
    
    target.Value = ""forgot something?""

End Sub";
        }
        public static string IgnoreQuickFixWorks_Expected()
        {
            return
@"
Private Sub Workbook_Open()
    
    Dim target As Range
'@Ignore ObjectVariableNotSet
    target = Range(""A1"")
    
    target.Value = ""forgot something?""

End Sub";
        }
        public static KeyValuePair<string, int> ForFunctionAssignment_ReturnsResult_TestParams()
        {
            var expectedResultCount = 2;
            var input =
@"
Private Function CombineRanges(ByVal source As Range, ByVal toCombine As Range) As Range
    If source Is Nothing Then
        CombineRanges = toCombine 'no inspection result (but there should be one!)
    Else
        CombineRanges = Union(source, toCombine) 'no inspection result (but there should be one!)
    End If
End Function";
            return new KeyValuePair<string, int>(input, expectedResultCount);
        }
        public static string ForFunctionAssignment_ReturnsResult_Expected()
        {
            return
@"
Private Function CombineRanges(ByVal source As Range, ByVal toCombine As Range) As Range
    If source Is Nothing Then
        Set CombineRanges = toCombine 'no inspection result (but there should be one!)
    Else
        Set CombineRanges = Union(source, toCombine) 'no inspection result (but there should be one!)
    End If
End Function";
        }
        public static KeyValuePair<string, int> ForPropertyGetAssignment_ReturnsResults_TestParams()
        {
            var expectedResultCount = 1;
            var input = @"
Private example As MyObject
Public Property Get Example() As MyObject
    Example = example
End Property
";
            return new KeyValuePair<string, int>(input, expectedResultCount);
        }
        public static string ForPropertyGetAssignment_ReturnsResults_Expected()
        {
            return
@"
Private example As MyObject
Public Property Get Example() As MyObject
    Set Example = example
End Property
";
        }
        public static KeyValuePair<string, int> LongPtrVariable_ReturnsNoResult_TestParams()
        {
            var expectResultCount = 0;
            var input =
@"
Private Sub TestLongPtr()
    Dim handle as LongPtr
    handle = 123456
End Sub";
            return new KeyValuePair<string, int>(input, expectResultCount);
        }
        public static KeyValuePair<string, int> NoTypeSpecified_ReturnsResult_TestParams()
        {
            var expectResultCount = 1;
            var input =
@"
Private Sub TestNonTyped(ByRef arg1)
    arg1 = new Collection
End Sub";
            return new KeyValuePair<string, int>(input, expectResultCount);
        }

        public static KeyValuePair<string, int> SelfAssigned_ReturnsNoResult_TestParams()
        {
            var expectResultCount = 0;
            var input =
@"
Private Sub TestSelfAssigned()
    Dim arg1 As new Collection
    arg1.Add 7
End Sub";
            return new KeyValuePair<string, int>(input, expectResultCount);
        }

        public static KeyValuePair<string, int> EnumVariable_ReturnsNoResult_TestParams()
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
            return new KeyValuePair<string, int>(input, expectResultCount);
        }
        public static KeyValuePair<string, int> FunctionReturnNotSet_ReturnsResult_TestParams()
        {
            var expectResultCount = 1;
            var input =
@"
Private Function TestFunctionReturn( start As Collection ) As Collection
    start.Add 5
    TestFunctionReturn = start
End Function";
            return new KeyValuePair<string, int>(input, expectResultCount);
        }

    }
}
