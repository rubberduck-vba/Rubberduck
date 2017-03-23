using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RubberduckTests.Inspections
{
    public static class AssignedByValParameterMakeLocalCopyQuickFixTests_VBABlocks
    {
        public static string LocalVariableAssignment_Input()
        {
            return
@"Public Sub Foo(ByVal arg1 As String)
    Let arg1 = ""test""
End Sub";
        }
        public static string LocalVariableAssignment_Expected()
        {
            return
@"Public Sub Foo(ByVal arg1 As String)
Dim localArg1 As String
localArg1 = arg1
    Let localArg1 = ""test""
End Sub";

        }
        public static string LocalVariableAssignment_ComplexFormat_Input()
        {
            return
@"Sub DoSomething(_
    ByVal foo As Long, _
    ByRef _
        bar, _
    ByRef barbecue _
                    )
    foo = 4
    bar = barbecue * _
                bar + foo / barbecue
End Sub
";

        }
        public static string LocalVariableAssignment_ComplexFormat_Expected()
        {
            return
@"Sub DoSomething(_
    ByVal foo As Long, _
    ByRef _
        bar, _
    ByRef barbecue _
                    )
Dim localFoo As Long
localFoo = foo
    localFoo = 4
    bar = barbecue * _
                bar + localFoo / barbecue
End Sub
";

        }
        public static string LocalVariableAssignment_ComputedNameAvoidsCollision_Input()
        {
            return
@"
Public Sub Foo(ByVal arg1 As String)
    Dim fooVar, _
        localArg1 As Long
    Let arg1 = ""test""
End Sub"
;

        }
        public static string LocalVariableAssignment_ComputedNameAvoidsCollision_Expected()
        {
            return
@"
Public Sub Foo(ByVal arg1 As String)
Dim localArg12 As String
localArg12 = arg1
    Dim fooVar, _
        localArg1 As Long
    Let localArg12 = ""test""
End Sub"
;
        }

        public static string[] LocalVariable_NameInUseOtherSub_SplitToken()
        {
            string[] splitToken = { "'VerifyNoChangeBelowThisLine" };
            return splitToken;
        }
        public static string LocalVariableAssignment_NameInUseOtherSub_Input()
        {
            return
@"
Public Function Bar2(ByVal arg2 As String) As String
    Dim arg1 As String
    Let arg1 = ""Test1""
    Bar2 = arg1
End Function

Public Sub Foo(ByVal arg1 As String)
    Let arg1 = ""test""
End Sub

'VerifyNoChangeBelowThisLine
Public Sub Bar(ByVal arg2 As String)
    Dim arg1 As String
    Let arg1 = ""Test2""
End Sub"
;
        }

        public static string LocalVariableAssignment_NameInUseOtherSub_Expected()
        {
            var inputCode = LocalVariableAssignment_NameInUseOtherSub_Input();
            string[] splitToken = LocalVariable_NameInUseOtherSub_SplitToken();
            return inputCode.Split(splitToken, System.StringSplitOptions.None)[1];
        }

        public static string LocalVariableAssignment_NameInUseOtherProperty_Input()
        {
            return
@"
Option Explicit
Private mBar as Long
Public Property Let Foo(ByVal bar As Long)
    bar = 42
    bar = bar * 2
    mBar = bar
End Property

Public Property Get Foo() As Long
    Dim bar as Long
    bar = 12
    Foo = mBar
End Property

'VerifyNoChangeBelowThisLine
Public Function bar() As Long
    Dim localBar As Long
    localBar = 7
    bar = localBar
End Function
";

        }

        public static string LocalVariableAssignment_NameInUseOtherProperty_Expected()
        {
            var inputCode = LocalVariableAssignment_NameInUseOtherProperty_Input();
            string[] splitToken = LocalVariable_NameInUseOtherProperty_SplitToken();
            return inputCode.Split(splitToken, System.StringSplitOptions.None)[1];
        }
        public static string[] LocalVariable_NameInUseOtherProperty_SplitToken()
        {
            string[] splitToken = { "'VerifyNoChangeBelowThisLine" };
            return splitToken;
        }

        public static string LocalVariableAssignment_UsesSet_Input()
        {
            return
@"
Public Sub Foo(FirstArg As Long, ByVal arg1 As Range)
    arg1 = Range(""A1: C4"")
End Sub"
;

        }

        public static string LocalVariableAssignment_UsesSet_Expected()
        {
            return
@"
Public Sub Foo(FirstArg As Long, ByVal arg1 As Range)
Dim localArg1 As Range
Set localArg1 = arg1
    localArg1 = Range(""A1: C4"")
End Sub"
;

        }

        public static string LocalVariableAssignment_NoAsTypeClause_Input()
        {
            return
            @"
Public Sub Foo(FirstArg As Long, ByVal arg1)
    arg1 = Range(""A1: C4"")
End Sub"
            ;
        }
        public static string LocalVariableAssignment_NoAsTypeClause_Expected()
        {
            return
            @"
Public Sub Foo(FirstArg As Long, ByVal arg1)
Dim localArg1 As Variant
localArg1 = arg1
    localArg1 = Range(""A1: C4"")
End Sub"
            ;
        }

        public static string LocalVariableAssignment_EnumType_Input()
        {
            return
            @"
Enum TestEnum
    EnumOne
    EnumTwo
    EnumThree
End Enum

Public Sub Foo(FirstArg As Long, ByVal arg1 As TestEnum)
    arg1 = EnumThree
End Sub"
            ;

        }
        public static string LocalVariableAssignment_EnumType_Expected()
        {
            return
            @"
Enum TestEnum
    EnumOne
    EnumTwo
    EnumThree
End Enum

Public Sub Foo(FirstArg As Long, ByVal arg1 As TestEnum)
Dim localArg1 As TestEnum
localArg1 = arg1
    localArg1 = EnumThree
End Sub"
            ;

        }
    }
}
