using System;
using Rubberduck.Resources.UnitTesting;

namespace Rubberduck.UnitTesting.CodeGeneration
{
    public partial class TestCodeGenerator
    {
        // These have to match the names of the exposed classes on the RD interface. They're only public here to facilitate a unit test that
        // ensures this (and hey - they're constants so it's not hurting anything, right?).
        public const string AssertClassName = "AssertClass";
        public const string PermissiveAssertClassName = "PermissiveAssertClass";

        private static string TestModuleBaseName => TestExplorer.UnitTest_NewModule_BaseName;
        private static string TestMethodBaseName => TestExplorer.UnitTest_NewMethod_BaseName;

        private static string DefaultTestCategory => TestExplorer.TestExplorer_Uncategorized;
        private static string RenameTestTodoComment => $"'TODO {TestExplorer.UnitTest_NewMethod_Rename}";
        private static string TestFailLabel => TestExplorer.UnitTest_NewMethod_TestFailLabel;
        private static string TestExitLabel => TestExplorer.UnitTest_NewMethod_TestExitLabel;
        private static string ArrangeComment => $"'{TestExplorer.UnitTest_NewMethod_Arrange}:";
        private static string ActComment => $"'{TestExplorer.UnitTest_NewMethod_Act}:";
        private static string AssertLabel => TestExplorer.UnitTest_NewMethod_Assert;
        private static string TestErrorMessage => $"{TestExplorer.UnitTest_NewMethod_RaisedTestError}: #";
        private static string ExpectedErrorConstant => TestExplorer.UnitTest_NewMethod_ExpectedError;
        private static string ExpectedErrorTodoComment => $"'TODO {TestExplorer.UnitTest_NewMethod_ChangeErrorNo}";
        private static string ExpectedErrorFailMessage => TestExplorer.UnitTest_NewMethod_ErrorNotRaised;

        private static string AccessCompareOption => $"Option Compare Database{Environment.NewLine}";
        private static string LateBindConstName => TestExplorer.UnitTest_NewModule_LateBindConstant;
        private static string LateBindDirectiveName => TestExplorer.UnitTest_NewModule_LateBindDirective;
        private static string DefaultTestFolder => TestExplorer.UnitTest_NewModule_DefaultFolder;
        private static string ModuleInitializeMethod => TestExplorer.UnitTest_NewMethod_ModuleInitializeMethod;
        private static string ModuleInitializeComment => $"'{TestExplorer.UnitTest_NewModule_RunOnce}.";
        private static string ModuleCleanupMethod => TestExplorer.UnitTest_NewMethod_ModuleCleanupMethod;
        private static string ModuleCleanupComment => $"'{TestExplorer.UnitTest_NewModule_RunOnce}.";
        private static string TestInitializeMethod => TestExplorer.UnitTest_NewMethod_TestInitializeMethod;
        private static string TestInitializeComment => $"'{TestExplorer.UnitTest_NewModule_RunBeforeTest}.";
        private static string TestCleanupMethod => TestExplorer.UnitTest_NewMethod_TestCleanupMethod;
        private static string TestCleanupComment => $"'{TestExplorer.UnitTest_NewModule_RunAfterTest}.";

        private static string TestMethodTemplate =>
$@"'@TestMethod(""{DefaultTestCategory}"")
Private Sub {{0}}() {RenameTestTodoComment}
    On Error GoTo {TestFailLabel}
    
    {ArrangeComment}

    {ActComment}

    '{AssertLabel}:
    Assert.Succeed

{TestExitLabel}:
    Exit Sub
{TestFailLabel}:
    Assert.Fail ""{TestErrorMessage}"" & Err.Number & "" - "" & Err.Description
End Sub";

        private static string TestMethodExpectedErrorTemplate =>
$@"'@TestMethod(""{DefaultTestCategory}"")
Private Sub {{0}}() {RenameTestTodoComment}
    Const {ExpectedErrorConstant} As Long = 0 {ExpectedErrorTodoComment}
    On Error GoTo {TestFailLabel}
    
    {ArrangeComment}

    {ActComment}

{AssertLabel}:
    Assert.Fail ""{ExpectedErrorFailMessage}""

{TestExitLabel}:
    Exit Sub
{TestFailLabel}:
    If Err.Number = {ExpectedErrorConstant} Then
        Resume {TestExitLabel}
    Else
        Resume {AssertLabel}
    End If
End Sub";

        private static string LateBindingDeclarations =>
@"    Private Assert As Object
    Private Fakes As Object";

        private static string EarlyBindingDeclarations =>
@"    Private Assert As Rubberduck.{0}
    Private Fakes As Rubberduck.FakesProvider";

        private static string DualBindingDeclarations =>
$@"#Const {LateBindConstName} = {LateBindDirectiveName}

#If {LateBindConstName} Then
{LateBindingDeclarations}
#Else
{EarlyBindingDeclarations}
#End If";

        private static string LateBindingInitialization =>
@"    Set Assert = CreateObject(""Rubberduck.{0}"")
    Set Fakes = CreateObject(""Rubberduck.FakesProvider"")";

        private static string EarlyBindingInitialization =>
@"    Set Assert = New Rubberduck.{0}
    Set Fakes = New Rubberduck.FakesProvider";

        private static string DualBindingInitialization =>
$@"#If {LateBindConstName} Then
{LateBindingInitialization}
#Else
{EarlyBindingInitialization}
#End If";

        private static string TestModuleTemplate =>
$@"{{0}}Option Explicit
Option Private Module

'@TestModule
'@Folder(""{DefaultTestFolder}"")

{{1}}

'@ModuleInitialize
Private Sub {ModuleInitializeMethod}()
{ModuleInitializeComment}
{{2}}
End Sub

'@ModuleCleanup
Private Sub {ModuleCleanupMethod}()
    {ModuleCleanupComment}
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub {TestInitializeMethod}()
    {TestInitializeComment}
End Sub

'@TestCleanup
Private Sub {TestCleanupMethod}()
    {TestCleanupComment}
End Sub";
    }
}
