using System;

namespace RubberduckTests.CodeExplorer
{
    internal static class CodeExplorerTestCode
    {
        public const string TestSubName = "TestSub";
        public const string TestSubWithLineLabelName = "TestSubWithLineLabel";
        public const string LineLabelName = "TestLineLabel";
        public const string TestSubWithUnresolvedMemberName = "TestSubWithUnresolvedMember";
        public const string TestFunctionName = "TestFunction";
        public const string UndeclaredVariableName = "undeclared";
        public const string TestPropertyName = "TestProperty";
        public const string TestPropertyParameterName = "value";
        public const string TestTypeName = "TestType";
        public const string TestTypeMemberName = "TestTypeMember";
        public const string TestTypeMemberTwoName = "TestTypeMemberTwo";
        public const string TestEnumName = "TestEnum";
        public const string TestEnumMemberName = "TestEnumMember";
        public const string TestEnumMemberTwoName = "TestEnumMemberTwo";
        public const string TestConstantName = "TestConstant";
        public const string TestFieldName = "TestField";
        public const string TestLibraryFunctionName = "TestLibraryFunction";
        public const string TestLibraryProcedureName = "TestLibraryProcedure";
        public const string TestEventName = "TestEvent";

        public const string MemberEnclosedVariableName = "variable";
        public const string MemberEnclosedConstantName = "constant";

        public static string TestSub =>
$@"
Public Sub {TestSubName}()
    Dim {MemberEnclosedVariableName} As String
End Sub
";

        public static string TestSubWithLineLabel =>
$@"
Public Sub {TestSubWithLineLabelName}()
    On Error GoTo {LineLabelName}
{LineLabelName}:
End Sub
";

        public static string TestSubWithUnresolvedMember =>
$@"
Public Sub {TestSubWithUnresolvedMemberName}()
    With New {CodeExplorerTestSetup.TestClassName}
        .UnresolvedMember
    End With
End Sub
";

        public static string TestFunction =>
$@"
Public Function {TestFunctionName}() As Variant
    {UndeclaredVariableName} = 42
    {TestFunctionName} = {UndeclaredVariableName}
End Function
";

        // TODO: This causes a parser error in testing due to no host application.
        public static string TestFunctionWithBracketedExpression => string.Empty;
//@"
//Public Function TestSubWithBracketedExpression() As Variant
//    TestSubWithBracketedExpression = [A1] 
//End Function
//";

        public static string TestProperty =>
$@"
Public Property Get {TestPropertyName}() As Variant
    Const {MemberEnclosedConstantName} = 42
    {TestPropertyName} = {MemberEnclosedConstantName}
End Property
Public Property Let {TestPropertyName}({TestPropertyParameterName} As Variant)
End Property
Public Property Set {TestPropertyName}({TestPropertyParameterName} As Variant)
End Property
";

        public static string TestType =>
$@"
Public Type {TestTypeName}
    {TestTypeMemberName} As Long
    {TestTypeMemberTwoName} As Long
End Type
";

        public static string TestEnum =>
$@"
Public Enum {TestEnumName}
    {TestEnumMemberName}
    {TestEnumMemberTwoName}
End Enum
";
        public static string TestUserFormCode =>
@"
Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    MsgBox ""Double-clicked""
End Sub
";

        public static string TestConstant => $"Public Const {TestConstantName} = 42";

        public static string TestField => $"Public {TestFieldName} As Long";

        public static string TestLibraryFunction => $"Public Declare PtrSafe Function {TestLibraryFunctionName} Lib \"test.dll\" (parameter As Long) As Long";

        public static string TestLibraryProcedure => $"Public Declare PtrSafe Sub {TestLibraryProcedureName} Lib \"test.dll\" ()";

        public static string TestEvent => $"Public Event {TestEventName}(parameter As Object)";

        public static string TestClassCode => string.Join(Environment.NewLine, TestEvent, TestProperty, TestSubWithUnresolvedMember);

        public static string TestModuleCode => 
            string.Join(Environment.NewLine, TestLibraryFunction, TestLibraryProcedure, TestField, TestType, TestEnum, TestFunction, TestSubWithLineLabel);

        public static string TestDocumentCode => string.Join(Environment.NewLine, TestConstant, TestSub, TestFunctionWithBracketedExpression);
    }
}
