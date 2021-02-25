using System;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using Moq;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;
using Rubberduck.Refactorings.EncapsulateFieldUseBackingField;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace RubberduckTests.Refactoring.EncapsulateField
{
    [TestFixture]
    public class EncapsulateFieldReferenceReplacerUDTFieldTests
    {
        [TestCase(true)]
        [TestCase(false)]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldReferenceReplacer))]
        public void PublicUDTField_ExternalReference(bool wrapInPrivateUDT)
        {
            var target = "targetField";
            var propertyName = "MyProperty";
            var testTargetTuple = (target, propertyName, false);

            var testModuleName = MockVbeBuilder.TestModuleName;
            var referenceExpression = $"{testModuleName}.{target}";
            var testModuleCode =
$@"
Option Explicit

Public Type TestType
    Fizz As Long
End Type

Public targetField As TestType";

            var declaringModule = (testModuleName, testModuleCode, ComponentType.StandardModule);


            var procedureModuleReferencingCode =
$@"
Option Explicit

Public Sub Bar()
    {referenceExpression}.Fizz = 7
End Sub
";
            var referencingModuleStdModule = (moduleName: "StdModule", procedureModuleReferencingCode, ComponentType.StandardModule);

            var refactoredCode = TestReferenceReplacement(wrapInPrivateUDT, testTargetTuple, declaringModule, referencingModuleStdModule);

            var referencingModuleCode = refactoredCode[referencingModuleStdModule.moduleName];

            StringAssert.Contains($"{testModuleName}.{propertyName}.Fizz = ", referencingModuleCode);
        }

        [TestCase(true)]
        [TestCase(false)]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldReferenceReplacer))]
        public void UDTField_PublicType_StdModuleReferenceWithMemberAccess(bool wrapInPrivateUDT)
        {
            var target = "targetField";
            var propertyName = "MyProperty";
            var testTargetTuple = (target, propertyName, false);

            var testModuleName = MockVbeBuilder.TestModuleName;

            var testModuleCode =
$@"
Public Type TBar
    First As String
    Second As Long
End Type

Public targetField As TBar";

            var declaringModule = (testModuleName, testModuleCode, ComponentType.StandardModule);

            var moduleReferencingCode =
$@"Option Explicit

'StdModule referencing the UDT

Public Sub FooBar()
    With {testModuleName}
        .targetField.First = ""Foo""
        .targetField.Second = 7
    End With
End Sub
";
            var referencingModuleStdModule = (moduleName: "StdModule", moduleReferencingCode, ComponentType.StandardModule);

            var refactoredCode = TestReferenceReplacement(wrapInPrivateUDT, testTargetTuple, declaringModule, referencingModuleStdModule);

            var referencingModuleCode = refactoredCode[referencingModuleStdModule.moduleName];

            StringAssert.Contains($"  .{propertyName}.First = ", referencingModuleCode);
            StringAssert.Contains($"  .{propertyName}.Second = ", referencingModuleCode);
        }

        [TestCase(true)]
        [TestCase(false)]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldReferenceReplacer))]
        public void UDTFieldSelection_ClassModuleSource_ExternalReference(bool wrapInPrivateUDT)
        {
            var target = "targetField";
            var propertyName = "MyProperty";
            var testTargetTuple = (target, propertyName, false);

            var testModuleName = MockVbeBuilder.TestModuleName;
            var classInstanceName = "theClass";
            var testModuleCode =
$@"
Option Explicit

Public targetField As TBar";
            var declaringModule = (testModuleName, testModuleCode, ComponentType.ClassModule);

            var moduleReferencingCode =
$@"
Option Explicit

Public Type TBar
    First As String
    Second As Long
End Type

Private {classInstanceName} As {testModuleName}

Public Sub Initialize()
    Set {classInstanceName} = New {testModuleName}
End Sub

Public Sub Fizz()
    {classInstanceName}.targetField.First = ""Foo""
End Sub

Public Sub Bang()
    {classInstanceName}.targetField.Second = 7
End Sub

Public Sub FizzBang()
    With {classInstanceName}
        .targetField.First = ""FizzBang""
        .targetField.Second = 7
    End With
End Sub
";

            var referencingModuleStdModule = (moduleName: "StdModule", moduleReferencingCode, ComponentType.StandardModule);
            var refactoredCode = TestReferenceReplacement(wrapInPrivateUDT, testTargetTuple, declaringModule, referencingModuleStdModule);

            var referencingModuleCode = refactoredCode[referencingModuleStdModule.moduleName];

            StringAssert.Contains($"{classInstanceName}.{propertyName}.First = ", referencingModuleCode);
            StringAssert.Contains($"{classInstanceName}.{propertyName}.Second = ", referencingModuleCode);
            StringAssert.Contains($"  .{propertyName}.Second = ", referencingModuleCode);
        }

        [TestCase(true, true)]
        [TestCase(true, false)]
        [TestCase(false, true)]
        [TestCase(false, false)]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldReferenceReplacer))]
        public void ReplaceUdtMemberReferences(bool wrapInPrivateUDT, bool isReadOnly)
        {
            var target = "myBazz";
            var udtMemberName = "MyBazz";

            var testTargetTuple = (target, udtMemberName, isReadOnly);

            var testModuleName = MockVbeBuilder.TestModuleName;
            var testModuleCode =
    $@"
Private Type TBazz
    FirstValue As String
    SecondValue As Long
End Type

Private myBazz As TBazz

Public Sub Fizz(newValue As String)
    myBazz.FirstValue = newValue
End Sub

Public Sub Bazz(newValue As Long)
    myBazz.SecondValue = newValue
End Sub
";
            var declaringModule = (testModuleName, testModuleCode, ComponentType.StandardModule);

            var refactoredCode = TestReferenceReplacement(wrapInPrivateUDT, testTargetTuple, declaringModule);

            var assignExpr_First = "  FirstValue = newValue";
            var assignExpr_Second = "  SecondValue = newValue";
            if (isReadOnly)
            {
                assignExpr_First = wrapInPrivateUDT ? $"  this.{udtMemberName}.FirstValue = newValue" : $"  {target}.FirstValue = newValue";
                assignExpr_Second = wrapInPrivateUDT ? $"  this.{udtMemberName}.SecondValue = newValue" : $"  {target}.SecondValue = newValue";
            }

            StringAssert.Contains(assignExpr_First, refactoredCode[MockVbeBuilder.TestModuleName]);
            StringAssert.Contains(assignExpr_Second, refactoredCode[MockVbeBuilder.TestModuleName]);
        }

        [TestCase(true, true)]
        [TestCase(true, false)]
        [TestCase(false, true)]
        [TestCase(false, false)]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldReferenceReplacer))]
        public void RenameFieldReferences_WithMemberAccess_NoExternalReferences(bool wrapInPrivateUDT, bool isReadOnly)
        {
            var target = "myBazz";
            var propertyName = "MyBazz";

            var testTargetTuple = (target, propertyName, isReadOnly);

            var testModuleName = MockVbeBuilder.TestModuleName;
            var testModuleCode =
    $@"
Private Type TBazz
    FirstValue As String
    SecondValue As Long
End Type

Private myBazz As TBazz

Public Sub Fizz(newValue As String)
    With myBazz
        .FirstValue = newValue
    End With
End Sub

Public Sub Bazz(newValue As String)
    With myBazz
        .SecondValue = newValue
    End With
End Sub
";

            var declaringModule = (testModuleName, testModuleCode, ComponentType.StandardModule);

            var refactoredCode = TestReferenceReplacement(wrapInPrivateUDT, testTargetTuple, declaringModule);

            var expectedWithStmt = wrapInPrivateUDT ? $"  With this.MyBazz{Environment.NewLine}" : $"  With myBazz{Environment.NewLine}";
            StringAssert.Contains(expectedWithStmt, refactoredCode[MockVbeBuilder.TestModuleName]);

            if (isReadOnly) //Get generated
            {
                StringAssert.Contains("  .FirstValue = newValue", refactoredCode[MockVbeBuilder.TestModuleName]);
                StringAssert.Contains("  .SecondValue = newValue", refactoredCode[MockVbeBuilder.TestModuleName]);
            }
            else //Let and Get generated
            {
                //The EF refactoring will create a FirstValue and SecondValue property - so the with member access expression
                //is replaced with the Let property name. The EF refactoring does not remove the 'With' statement block even 
                //though it is no longer required by this specific scenario
                StringAssert.Contains("  FirstValue = newValue", refactoredCode[MockVbeBuilder.TestModuleName]);
                StringAssert.Contains("  SecondValue = newValue", refactoredCode[MockVbeBuilder.TestModuleName]);
            }
        }

        [TestCase(true, true)]
        [TestCase(true, false)]
        [TestCase(false, true)]
        [TestCase(false, false)]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldReferenceReplacer))]
        public void ReplaceAccessorExpression(bool wrapInPrivateUDT, bool isReadOnly)
        {
            var target = "myBazz";
            var udtMemberName = "MyBazz";

            var testTargetTuple = (target, udtMemberName, isReadOnly);

            var testModuleName = MockVbeBuilder.TestModuleName;
            var testModuleCode =
    $@"
Private Type TBazz
    FirstValue As String
    SecondValue As Long
End Type

Private myBazz As TBazz

Public Function GetTheFirstValue() As String
    GetTheFirstValue = myBazz.FirstValue
End Function

Public Sub SetTheFirstValue(arg As Long)
    myBazz.FirstValue = arg
End Sub
";

            var declaringModule = (testModuleName, testModuleCode, ComponentType.StandardModule);

            var refactoredCode = TestReferenceReplacement(wrapInPrivateUDT, testTargetTuple, declaringModule);

            var firstValueWrite = "FirstValue = arg";
            if (isReadOnly)
            {
                firstValueWrite = wrapInPrivateUDT ? $"this.{udtMemberName}.FirstValue = arg" : "myBazz.FirstValue = arg";
            }

            StringAssert.Contains(firstValueWrite, refactoredCode[MockVbeBuilder.TestModuleName]);
            StringAssert.Contains("GetTheFirstValue = FirstValue", refactoredCode[MockVbeBuilder.TestModuleName]);
        }

        [TestCase(true)]
        [TestCase(false)]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldReferenceReplacer))]
        public void ReplacePublicUDTAccessorExpressions(bool wrapInPrivateUDT)
        {
            var target = "myBazz";
            var propertyName = "MyProperty";

            var testTargetTuple = (target, propertyName, false);

            var testModuleName = MockVbeBuilder.TestModuleName;
            var testModuleCode =
    $@"

Public Type TBazz
    FirstValue As String
    SecondValue As Long
End Type

Public myBazz As TBazz

'Public Function GetTheFirstValue() As String
'    GetTheFirstValue = myBazz.FirstValue
'End Function

'Public Sub SetTheFirstValue(arg As Long)
'    myBazz.FirstValue = arg
'End Sub
";
            var declaringModule = (testModuleName, testModuleCode, ComponentType.StandardModule);

            var referencingModule = "AnotherModule";
            var referencingModuleCode =
    $@"

Public Function GetBazzFirst() As String
    GetBazzFirst = myBazz.FirstValue
End Function

Public Sub SetBazzFirst(arg As Long)
    myBazz.FirstValue = arg
End Sub
";

            var referencingModuleStdModule = (moduleName: referencingModule, referencingModuleCode, ComponentType.StandardModule);

            var refactoredCode = TestReferenceReplacement(wrapInPrivateUDT, testTargetTuple, declaringModule, referencingModuleStdModule);

            StringAssert.Contains($"GetBazzFirst = {MockVbeBuilder.TestModuleName}.{propertyName}.FirstValue", refactoredCode[referencingModule]);
            StringAssert.Contains($"{MockVbeBuilder.TestModuleName}.{propertyName}.FirstValue = arg", refactoredCode[referencingModule]);
        }

        [TestCase(true, true)]
        [TestCase(true, false)]
        [TestCase(false, true)]
        [TestCase(false, false)]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldReferenceReplacer))]
        public void ModifiesCorrectUDTMemberReferences_MemberAccess(bool wrapInPrivateUDT, bool isReadOnly)
        {
            var target = "targetField";
            var udtMemberName = "TargetField";

            var testTargetTuple = (target, udtMemberName, isReadOnly);

            var testModuleName = MockVbeBuilder.TestModuleName;
            var testModuleCode =
    $@"
Private Type TBar
    First As String
    Second As Long
End Type

Private targetField As TBar

Private bogeyField As TBar

Public Sub Foo(arg1 As String, arg2 As Long)
    targetField.First = arg1
    bogeyField.First = arg1
    targetField.Second = arg2
    bogeyField.Second = arg2
End Sub
";
            var declaringModule = (testModuleName, testModuleCode, ComponentType.StandardModule);

            var refactoredCode = TestReferenceReplacement(wrapInPrivateUDT, testTargetTuple, declaringModule);
            var actualCode = refactoredCode[testModuleName];

            var targetFirstAssignment = $" First = arg1";
            var targetSecondAssignment = $" Second = arg2";

            if (isReadOnly)
            {
                targetFirstAssignment = wrapInPrivateUDT ? $" this.{udtMemberName}.First = arg1" : $" {target}.First = arg1";
                targetSecondAssignment = wrapInPrivateUDT ? $" this.{udtMemberName}.Second = arg2" : $" {target}.Second = arg2";
            }

            StringAssert.Contains(targetFirstAssignment, actualCode);
            StringAssert.Contains(targetSecondAssignment, actualCode);

            StringAssert.Contains($"bogeyField.First = arg1", actualCode);
            StringAssert.Contains($"bogeyField.Second = arg2", actualCode);
        }

        [TestCase(true, true)]
        [TestCase(true, false)]
        [TestCase(false, true)]
        [TestCase(false, false)]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldReferenceReplacer))]
        public void NestedUDTMembers(bool wrapInPrivateUDT, bool isReadOnly)
        {
            var target = "mTypesField";
            var udtMemberName = "TypesField";

            var testTargetTuple = (target, udtMemberName, isReadOnly);

            var testModuleName = MockVbeBuilder.TestModuleName;
            var testModuleCode =
    $@"
Option Explicit

Private Type PType1
    FirstValType1 As Long
    SecondValType1 As String
End Type

Private Type PType2
    FirstValType2 As Long
    SecondValType2 As String
    Third As PType1
End Type

Private mTypesField As PType2

Private Sub Class_Initialize()
    mTypesField.Third.SecondValType1 = ""Wah""
End Sub

Private Sub TestSub2()
    TestSub3 mTypesField.Third.SecondValType1
End Sub

Private Sub TestSub3(ByVal arg As String)
End Sub
";

            var declaringModule = (testModuleName, testModuleCode, ComponentType.StandardModule);

            var refactoredCode = TestReferenceReplacement(wrapInPrivateUDT, testTargetTuple, declaringModule);

            var expectedAssignment = "SecondValType1 = \"Wah\"";
            if (isReadOnly)
            {
                expectedAssignment = wrapInPrivateUDT ? $"this.{udtMemberName}.Third.{expectedAssignment}" : "TypesField.Third.SecondValType1 = \"Wah\"";
            }
            StringAssert.Contains(expectedAssignment, refactoredCode[testModuleName]);

            StringAssert.Contains("TestSub3 SecondValType1", refactoredCode[testModuleName]);
        }

        [TestCase(true)]
        [TestCase(false)]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldReferenceReplacer))]
        public void PublicUDTField_ExternalRefNestedWithStatement(bool wrapInPrivateUDT)
        {
            var target = "mTypesField";

            var testTargetTuple = (target, "TypesField", false);

            var testModuleName = MockVbeBuilder.TestModuleName;
            var testModuleCode =
    $@"
Option Explicit

Public Type PType1
    FirstValType1 As Long
End Type

Public Type PType2
    Third As PType1
End Type

Public mTypesField As PType2
";

            var declaringModule = (testModuleName, testModuleCode, ComponentType.StandardModule);

            var referencingModuleName = "AnotherModule";
            var referencingModuleCode =
    $@"
Private testVal As Long

Public Sub TestSub()
    With {testModuleName}
        With .mTypesField
            With .Third
                testVal = .FirstValType1
            End With
        End With
    End With
End Sub
";

            var referencingModuleStdModule = (moduleName: referencingModuleName, referencingModuleCode, ComponentType.StandardModule);

            var refactoredCode = TestReferenceReplacement(wrapInPrivateUDT, testTargetTuple, declaringModule, referencingModuleStdModule);

            StringAssert.Contains($"With .TypesField", refactoredCode[referencingModuleName]);

            StringAssert.Contains(" testVal = .FirstValType1", refactoredCode[referencingModuleName]);
        }

        [TestCase(true, true)]
        [TestCase(true, false)]
        [TestCase(false, true)]
        [TestCase(false, false)]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldReferenceReplacer))]
        public void PrivateUDTField_RefNestedWithStatements(bool wrapInPrivateUDT, bool isReadOnly)
        {
            var target = "mTypesField";
            var udtMemberName = "TypesField";


            var testTargetTuple = (target, udtMemberName, isReadOnly);

            var testModuleName = MockVbeBuilder.TestModuleName;
            var testModuleCode =
    $@"
Option Explicit

Private Type PType1
    FirstValType1 As Long
End Type

Private Type PType2
    Third As PType1
End Type

Private mTypesField As PType2

Public Sub TestSub(ByVal arg As Long)
    With mTypesField
        With .Third
            .FirstValType1 = arg
        End With
    End With
End Sub

Public Function TestFunc() As Long
    With mTypesField
        With .Third
            TestFunc = .FirstValType1
        End With
    End With
End Function

";

            var declaringModule = (testModuleName, testModuleCode, ComponentType.StandardModule);

            var refactoredCode = TestReferenceReplacement(wrapInPrivateUDT, testTargetTuple, declaringModule);

            var typesFieldWithStmt = wrapInPrivateUDT ? $"With this.{udtMemberName}" : $"With {target}";
            StringAssert.Contains(typesFieldWithStmt, refactoredCode[testModuleName]);

            StringAssert.Contains($"With .Third", refactoredCode[testModuleName]);

            var expectedAssignment = isReadOnly ? ".FirstValType1 = arg" : "FirstValType1 = arg";

            StringAssert.Contains(expectedAssignment, refactoredCode[testModuleName]);

            StringAssert.Contains("TestFunc = FirstValType1", refactoredCode[testModuleName]);
        }

        [TestCase(true, true)]
        [TestCase(true, false)]
        [TestCase(false, true)]
        [TestCase(false, false)]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldReferenceReplacer))]
        public void PrivateUDTFieldMultipleMembers(bool wrapInPrivateUDT, bool isReadOnly)
        {
            var target = "mVehicle";
            var wrappedUDTMemberName = "Vehicle";

            var testTargetTuple = (target, wrappedUDTMemberName, isReadOnly);

            var testModuleName = MockVbeBuilder.TestModuleName;
            var testModuleCode =
    $@"
Option Explicit

Private Type TVehicle
    Seats As Integer
    Wheels As Integer
End Type

Private mVehicle As TVehicle

Private Sub Class_Initialize()
    mVehicle.Wheels = 4
    mVehicle.Seats = 2
End Sub
";

            var declaringModule = (testModuleName, testModuleCode, ComponentType.StandardModule);
            var refactoredCode = TestReferenceReplacement(wrapInPrivateUDT, testTargetTuple, declaringModule);
            var result = refactoredCode[MockVbeBuilder.TestModuleName];

            var expectedAssignmentExpressionWheels = " Wheels = 4";
            if (isReadOnly)
            {
                expectedAssignmentExpressionWheels = wrapInPrivateUDT ? $" this.{wrappedUDTMemberName}.Wheels = 4" : " mVehicle.Wheels = 4";
            }

            var expectedAssignmentExpressionSeats = " Seats = 2";
            if (isReadOnly)
            {
                expectedAssignmentExpressionSeats = wrapInPrivateUDT ? $" this.{wrappedUDTMemberName}.Seats = 2" : " mVehicle.Seats = 2";
            }

            StringAssert.Contains(expectedAssignmentExpressionWheels, result);
            StringAssert.Contains(expectedAssignmentExpressionSeats, result);
        }
        private static IDictionary<string, string> TestReferenceReplacement(bool wrapInPrivateUDT, (string, string, bool) testTargetTuple, params (string, string, ComponentType)[] moduleTuples)
        {
            return ReferenceReplacerTestSupport.TestReferenceReplacement(wrapInPrivateUDT, testTargetTuple, moduleTuples);
        }
    }
}
