using NUnit.Framework;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.Utility;
using System.Collections.Generic;

namespace RubberduckTests.Refactoring.EncapsulateField
{
    [TestFixture]
    public class EncapsulatedUDTFieldTests : InteractiveRefactoringTestBase<IEncapsulateFieldPresenter, EncapsulateFieldModel>
    {
        private EncapsulateFieldTestSupport Support { get; } = new EncapsulateFieldTestSupport();

        [TestCase("Public")]
        [TestCase("Private")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UserDefinedType_UserAcceptsDefaults(string accessibility)
        {
            string inputCode =
$@"
Private Type TBar
    First As String
    Second As Long
End Type

{accessibility} th|is As TBar";


            var presenterAction = Support.UserAcceptsDefaults();

            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            StringAssert.Contains("Private this As TBar", actualCode);
            StringAssert.DoesNotContain("this = value", actualCode);
            StringAssert.DoesNotContain($"This = this", actualCode);
            StringAssert.Contains($"Public Property Get First", actualCode);
            StringAssert.Contains($"Public Property Get Second", actualCode);
            StringAssert.Contains($"this.First = {Support.RHSIdentifier}", actualCode);
            StringAssert.Contains($"First = this.First", actualCode);
            StringAssert.Contains($"this.Second = {Support.RHSIdentifier}", actualCode);
            StringAssert.Contains($"Second = this.Second", actualCode);
        }

        [TestCase(true, true)]
        [TestCase(false, false)]
        [TestCase(true, false)]
        [TestCase(false, true)]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UserDefinedType_TwoFields(bool encapsulateThis, bool encapsulateThat)
        {
            string inputCode =
$@"
Private Type TBar
    First As String
    Second As Long
End Type

Public th|is As TBar
Public that As TBar";

            var validator = EncapsulateFieldValidationsProvider.NameOnlyValidator(NameValidators.Default);
            var expectedThis = new EncapsulationIdentifiers("this", validator);
            var expectedThat = new EncapsulationIdentifiers("that", validator);

            var userInput = new UserInputDataObject()
                    .AddUserInputSet(expectedThis.TargetFieldName, encapsulationFlag: encapsulateThis)
                    .AddUserInputSet(expectedThat.TargetFieldName, encapsulationFlag: encapsulateThat);

            var presenterAction = Support.SetParameters(userInput);

            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            if (encapsulateThis && encapsulateThat)
            {
                StringAssert.Contains($"Private {expectedThis.TargetFieldName} As TBar", actualCode);
                StringAssert.Contains($"First = {expectedThis.TargetFieldName}.First", actualCode);
                StringAssert.Contains($"Second = {expectedThis.TargetFieldName}.Second", actualCode);
                StringAssert.Contains($"{expectedThis.TargetFieldName}.First = {Support.RHSIdentifier}", actualCode);
                StringAssert.Contains($"{expectedThis.TargetFieldName}.Second = {Support.RHSIdentifier}", actualCode);
                StringAssert.Contains($"Property Get First", actualCode);
                StringAssert.Contains($"Property Get Second", actualCode);

                StringAssert.Contains($"Private {expectedThat.TargetFieldName} As TBar", actualCode);
                StringAssert.Contains($"First_1 = {expectedThat.TargetFieldName}.First", actualCode);
                StringAssert.Contains($"Second_1 = {expectedThat.TargetFieldName}.Second", actualCode);
                StringAssert.Contains($"{expectedThat.TargetFieldName}.First = {Support.RHSIdentifier}", actualCode);
                StringAssert.Contains($"{expectedThat.TargetFieldName}.Second = {Support.RHSIdentifier}", actualCode);
                StringAssert.Contains($"Property Get First_1", actualCode);
                StringAssert.Contains($"Property Get Second_1", actualCode);

                StringAssert.Contains($"Private {expectedThis.TargetFieldName} As TBar", actualCode);
                StringAssert.Contains($"Private {expectedThat.TargetFieldName} As TBar", actualCode);
            }
            else if (encapsulateThis && !encapsulateThat)
            {
                StringAssert.Contains($"First = {expectedThis.TargetFieldName}.First", actualCode);
                StringAssert.Contains($"Second = {expectedThis.TargetFieldName}.Second", actualCode);
                StringAssert.Contains($"{expectedThis.TargetFieldName}.First = {Support.RHSIdentifier}", actualCode);
                StringAssert.Contains($"{expectedThis.TargetFieldName}.Second = {Support.RHSIdentifier}", actualCode);
                StringAssert.Contains($"Property Get First", actualCode);
                StringAssert.Contains($"Property Get Second", actualCode);

                StringAssert.Contains($"Private {expectedThis.TargetFieldName} As TBar", actualCode);
                StringAssert.Contains($"Public {expectedThat.TargetFieldName} As TBar", actualCode);
            }
            else if (!encapsulateThis && encapsulateThat)
            {
                StringAssert.Contains($"First = {expectedThat.TargetFieldName}.First", actualCode);
                StringAssert.Contains($"Second = {expectedThat.TargetFieldName}.Second", actualCode);
                StringAssert.Contains($"{expectedThat.TargetFieldName}.First = {Support.RHSIdentifier}", actualCode);
                StringAssert.Contains($"{expectedThat.TargetFieldName}.Second = {Support.RHSIdentifier}", actualCode);
                StringAssert.Contains($"Property Get First", actualCode);
                StringAssert.Contains($"Property Get Second", actualCode);

                StringAssert.Contains($"Public {expectedThis.TargetFieldName} As TBar", actualCode);
                StringAssert.Contains($"Private {expectedThat.TargetFieldName} As TBar", actualCode);
            }
            else
            {
                StringAssert.Contains($"Public {expectedThis.TargetFieldName} As TBar", actualCode);
                StringAssert.Contains($"Public {expectedThat.TargetFieldName} As TBar", actualCode);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void ModifiesCorrectUDTMemberReference_MemberAccess()
        {
            string inputCode =
$@"
Private Type TBar
    First As String
    Second As Long
End Type

Private th|is As TBar

Private that As TBar

Public Sub Foo(arg1 As String, arg2 As Long)
    this.First = arg1
    that.First = arg1
    this.Second = arg2
    that.Second = arg2
End Sub
";


            var presenterAction = Support.UserAcceptsDefaults();

            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            StringAssert.Contains($"this.First = {Support.RHSIdentifier}", actualCode);
            StringAssert.Contains($"First = this.First", actualCode);
            StringAssert.Contains($"this.Second = {Support.RHSIdentifier}", actualCode);
            StringAssert.Contains($"Second = this.Second", actualCode);
            StringAssert.Contains($" First = arg1", actualCode);
            StringAssert.Contains($" Second = arg2", actualCode);
            StringAssert.Contains($"that.First = arg1", actualCode);
            StringAssert.Contains($"that.Second = arg2", actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void ModifiesCorrectUDTMemberReference_WithMemberAccess()
        {
            string inputCode =
$@"
Private Type TBar
    First As String
    Second As Long
End Type

Private th|is As TBar

Private that As TBar

Public Sub Foo(arg1 As String, arg2 As Long)
    With this
        .First = arg1
        .Second = arg2
    End With

    With that
        .First = arg1
        .Second = arg2
    End With
End Sub
";
            var presenterAction = Support.UserAcceptsDefaults();

            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            StringAssert.DoesNotContain($" First = arg1", actualCode);
            StringAssert.DoesNotContain($" Second = arg2", actualCode);
            StringAssert.Contains($" .First = arg1", actualCode);
            StringAssert.Contains($" .Second = arg2", actualCode);
            StringAssert.Contains("With this", actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void ModifiesCorrectUDTMemberReference_WithMemberAccessExternal()
        {
            string sourceModuleName = "SourceModule";
            string inputCode =
$@"
Public Type TBar
    First As String
    Second As Long
End Type

Public th|is As TBar

Private that As TBar
";
            string module2 =
$@"
Public Sub Foo(arg1 As String, arg2 As Long)
    With {sourceModuleName}.this
        .First = arg1
        .Second = arg2
    End With

    With that
        .First = arg1
        .Second = arg2
    End With
End Sub
";

            var presenterAction = Support.UserAcceptsDefaults();

            var codeString = inputCode.ToCodeString();

            var actualModuleCode = RefactoredCode(
                sourceModuleName,
                codeString.CaretPosition.ToOneBased(),
                presenterAction,
                null,
                false,
                ("Module2", module2, ComponentType.StandardModule),
                (sourceModuleName, codeString.Code, ComponentType.StandardModule));

            var actualCode = actualModuleCode["Module2"];
            var sourceCode = actualModuleCode[sourceModuleName];

            StringAssert.DoesNotContain($" First = arg1", actualCode);
            StringAssert.DoesNotContain($" Second = arg2", actualCode);
            StringAssert.Contains($" .First = arg1", actualCode);
            StringAssert.Contains($" .Second = arg2", actualCode);
            StringAssert.Contains($"With {sourceModuleName}.This", actualCode);
        }

        [TestCase("Public")]
        [TestCase("Private")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UserDefinedTypeMembersAndFields(string accessibility)
        {
            string inputCode =
$@"
Private Type TBar
    First As String
    Second As Long
End Type

Public mFoo As String
Public mBar As Long
Private mFizz

{accessibility} t|his As TBar";


            var userInput = new UserInputDataObject()
            .UserSelectsField("this", "MyType")
            .UserSelectsField("mFoo", "Foo")
            .UserSelectsField("mBar", "Bar")
            .UserSelectsField("mFizz", "Fizz");

            var presenterAction = Support.SetParameters(userInput);

            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);

            StringAssert.Contains("Private this As TBar", actualCode);
            StringAssert.DoesNotContain("this = value", actualCode);
            StringAssert.DoesNotContain("MyType = this", actualCode);
            StringAssert.Contains($"this.First = {Support.RHSIdentifier}", actualCode);
            StringAssert.Contains($"First = this.First", actualCode);
            StringAssert.Contains($"this.Second = {Support.RHSIdentifier}", actualCode);
            StringAssert.Contains($"Second = this.Second", actualCode);
            StringAssert.DoesNotContain($"Second = Second", actualCode);
            StringAssert.Contains($"Private mFoo As String", actualCode);
            StringAssert.Contains($"Public Property Get Foo(", actualCode);
            StringAssert.Contains($"Public Property Let Foo(", actualCode);
            StringAssert.Contains($"Private mBar As Long", actualCode);
            StringAssert.Contains($"Public Property Get Bar(", actualCode);
            StringAssert.Contains($"Public Property Let Bar(", actualCode);
            StringAssert.Contains($"Private mFizz As Variant", actualCode);
            StringAssert.Contains($"Public Property Get Fizz() As Variant", actualCode);
            StringAssert.Contains($"Public Property Let Fizz(", actualCode);
        }

        [TestCase("Public")]
        [TestCase("Private")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UserDefinedTypeMember_ContainsObjects(string accessibility)
        {
            string inputCode =
$@"
Private Type TBar
    First As Class1
    Second As Long
End Type

{accessibility} th|is As TBar";

            string class1Code =
@"Option Explicit

Public Sub Foo()
End Sub
";

            var userInput = new UserInputDataObject()
                .UserSelectsField("this", "MyType");

            var presenterAction = Support.SetParameters(userInput);

            var codeString = inputCode.ToCodeString();
            var actualModuleCode = RefactoredCode(
                "Module1",
                codeString.CaretPosition.ToOneBased(),
                presenterAction,
                null,
                false,
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Module1", codeString.Code, ComponentType.StandardModule));

            var actualCode = actualModuleCode["Module1"];

            StringAssert.Contains("Private this As TBar", actualCode);
            StringAssert.DoesNotContain($"this = {Support.RHSIdentifier}", actualCode);
            StringAssert.DoesNotContain("MyType = this", actualCode);
            StringAssert.Contains($"Property Set First(ByVal {Support.RHSIdentifier} As Class1)", actualCode);
            StringAssert.Contains("Property Get First() As Class1", actualCode);
            StringAssert.Contains($"Set this.First = {Support.RHSIdentifier}", actualCode);
            StringAssert.Contains($"Set First = this.First", actualCode);
            StringAssert.Contains($"this.Second = {Support.RHSIdentifier}", actualCode);
            StringAssert.Contains($"Second = this.Second", actualCode);
            StringAssert.DoesNotContain($"Second = Second", actualCode);
            Assert.AreEqual(actualCode.IndexOf($"this.First = {Support.RHSIdentifier}"), actualCode.LastIndexOf($"this.First = {Support.RHSIdentifier}"));
        }

        [TestCase("Public")]
        [TestCase("Private")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UserDefinedTypeMember_ContainsVariant(string accessibility)
        {
            string inputCode =
$@"
Private Type TBar
    First As String
    Second As Variant
End Type

{accessibility} th|is As TBar";

            var userInput = new UserInputDataObject()
                .UserSelectsField("this", "MyType");

            var presenterAction = Support.SetParameters(userInput);

            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            StringAssert.Contains("Private this As TBar", actualCode);
            StringAssert.DoesNotContain($"this = {Support.RHSIdentifier}", actualCode);
            StringAssert.DoesNotContain("MyType = this", actualCode);
            StringAssert.Contains($"this.First = {Support.RHSIdentifier}", actualCode);
            StringAssert.Contains($"First = this.First", actualCode);
            StringAssert.Contains($"IsObject", actualCode);
            StringAssert.Contains($"this.Second = {Support.RHSIdentifier}", actualCode);
            StringAssert.Contains($"Second = this.Second", actualCode);
            StringAssert.DoesNotContain($"Second = Second", actualCode);
        }

        [TestCase("Public")]
        [TestCase("Private")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UserDefinedTypeMember_ContainsArrays(string accessibility)
        {
            string inputCode =
$@"
Private Type TBar
    First(5) As String
    Second(1 to 100) As Long
    Third() As Double
End Type

{accessibility} t|his As TBar";

            var userInput = new UserInputDataObject()
                .UserSelectsField("this", "MyType");

            var presenterAction = Support.SetParameters(userInput);

            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            StringAssert.Contains($"Private this As TBar", actualCode);
            StringAssert.DoesNotContain($"this = {Support.RHSIdentifier}", actualCode);
            StringAssert.DoesNotContain($"this.First = {Support.RHSIdentifier}", actualCode);
            StringAssert.Contains($"Property Get First() As Variant", actualCode);
            StringAssert.Contains($"First = this.First", actualCode);
            StringAssert.DoesNotContain($"this.Second = {Support.RHSIdentifier}", actualCode);
            StringAssert.Contains($"Second = this.Second", actualCode);
            StringAssert.Contains($"Property Get Second() As Variant", actualCode);
            StringAssert.Contains($"Third = this.Third", actualCode);
            StringAssert.Contains($"Property Get Third() As Variant", actualCode);
        }

        [TestCase("Public")]
        [TestCase("Private")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UserDefinedTypeMembers_ExternallyDefinedType(string accessibility)
        {
            string inputCode =
$@"
Option Explicit

{accessibility} th|is As TBar";

            string typeDefinition =
$@"
Public Type TBar
    First As String
    Second As Long
End Type
";

            var userInput = new UserInputDataObject()
                .UserSelectsField("this", "MyType");

            var presenterAction = Support.SetParameters(userInput);

            var codeString = inputCode.ToCodeString();
            var actualModuleCode = RefactoredCode(
                "Class1",
                codeString.CaretPosition.ToOneBased(),
                presenterAction,
                null,
                false,
                ("Class1", codeString.Code, ComponentType.ClassModule),
                ("Module1", typeDefinition, ComponentType.StandardModule));

            Assert.AreEqual(typeDefinition, actualModuleCode["Module1"]);


            var actualCode = actualModuleCode["Class1"];
            StringAssert.Contains("Private this As TBar", actualCode);

            if (accessibility == "Public")
            {
                StringAssert.Contains($"this = {Support.RHSIdentifier}", actualCode);
                StringAssert.Contains("MyType = this", actualCode);
                StringAssert.Contains($"Public Property Get MyType", actualCode);
                Assert.AreEqual(actualCode.IndexOf("MyType = this"), actualCode.LastIndexOf("MyType = this"));
                StringAssert.DoesNotContain($"this.First = {Support.RHSIdentifier}", actualCode);
                StringAssert.DoesNotContain($"this.Second = {Support.RHSIdentifier}", actualCode);
                StringAssert.DoesNotContain($"Second = Second", actualCode);
                StringAssert.Contains($"Public Property Let MyType(ByRef {Support.RHSIdentifier} As TBar", actualCode);
            }
        }

        [TestCase("Public")]
        [TestCase("Private")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UserDefinedTypeMembers_ObjectField(string accessibility)
        {
            string inputCode =
$@"
Option Explicit

{accessibility} mThe|Class As Class1";

            string classContent =
$@"
Option Explicit

Private Sub Class_Initialize()
End Sub
";

            var presenterAction = Support.UserAcceptsDefaults();

            var codeString = inputCode.ToCodeString();
            var actualModuleCode = RefactoredCode(
                "Module1",
                codeString.CaretPosition.ToOneBased(),
                presenterAction,
                null,
                false,
                ("Class1", classContent, ComponentType.ClassModule),
                ("Module1", codeString.Code, ComponentType.StandardModule));

            var actualCode = actualModuleCode["Module1"];

            StringAssert.Contains($"Private mTheClass As Class1", actualCode);
            StringAssert.Contains($"Set mTheClass = {Support.RHSIdentifier}", actualCode);
            StringAssert.Contains($"Set TheClass = mTheClass", actualCode);
            StringAssert.Contains($"Public Property Set TheClass", actualCode);
        }

        [TestCase("SourceModule", "this")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void StdModuleSource_PublicUDTField_PublicType_StdModuleReference(string sourceModuleName, string referenceQualifier)
        {
            var sourceModuleCode =
$@"
Public Type TBar
    First As String
    Second As Long
End Type

Public th|is As TBar";

            var moduleReferencingCode =
$@"Option Explicit

'StdModule referencing the UDT

Private Const fooConst As String = ""Foo""

Private Const barConst As Long = 7

Public Sub Foo()
    {referenceQualifier}.First = fooConst
End Sub

Public Sub Bar()
    {referenceQualifier}.Second = barConst
End Sub

Public Sub FooBar()
    With {sourceModuleName}
        .this.First = fooConst
        .this.Second = barConst
    End With
End Sub
";
            var userInput = new UserInputDataObject()
                .UserSelectsField("this", "MyType");

            var presenterAction = Support.SetParameters(userInput);

            var sourceCodeString = sourceModuleCode.ToCodeString();

            var actualModuleCode = RefactoredCode(
                sourceModuleName,
                sourceCodeString.CaretPosition.ToOneBased(),
                presenterAction,
                null,
                false,
                ("StdModule", moduleReferencingCode, ComponentType.StandardModule),
                (sourceModuleName, sourceCodeString.Code, ComponentType.StandardModule));

            var referencingModuleCode = actualModuleCode["StdModule"];
            StringAssert.Contains($"{sourceModuleName}.MyType.First = ", referencingModuleCode);
            StringAssert.Contains($"{sourceModuleName}.MyType.Second = ", referencingModuleCode);
            StringAssert.Contains($"  .MyType.Second = ", referencingModuleCode);
        }

        private IDictionary<string,string> Scenario_StdModuleSource_StandardAndClassReferencingModules(string referenceQualifier, string typeDeclarationAccessibility, string sourceModuleName, UserInputDataObject userInput)
        {
            var sourceModuleCode =
$@"
{typeDeclarationAccessibility} Type TBar
    First As String
    Second As Long
End Type

Public th|is As TBar";

            var procedureModuleReferencingCode =
$@"Option Explicit

'StdModule referencing the UDT

Private Const foo As String = ""Foo""

Private Const bar As Long = 7

Public Sub Foo()
    {referenceQualifier}.First = foo
End Sub

Public Sub Bar()
    {referenceQualifier}.Second = bar
End Sub

Public Sub FooBar()
    With {sourceModuleName}
        .this.First = foo
        .this.Second = bar
    End With
End Sub
";

            string classModuleReferencingCode =
$@"Option Explicit

'ClassModule referencing the UDT

Private Const foo As String = ""Foo""

Private Const bar As Long = 7

Public Sub Foo()
    {referenceQualifier}.First = foo
End Sub

Public Sub Bar()
    {referenceQualifier}.Second = bar
End Sub

Public Sub FooBar()
    With {sourceModuleName}
        .this.First = foo
        .this.Second = bar
    End With
End Sub
";

            var presenterAction = Support.SetParameters(userInput);

            var sourceCodeString = sourceModuleCode.ToCodeString();

            //Only Public Types are accessible to ClassModules
            if (typeDeclarationAccessibility.Equals("Public"))
            {
                return RefactoredCode(
                    sourceModuleName,
                    sourceCodeString.CaretPosition.ToOneBased(),
                    presenterAction,
                    null,
                    false,
                    ("StdModule", procedureModuleReferencingCode, ComponentType.StandardModule),
                    ("ClassModule", classModuleReferencingCode, ComponentType.ClassModule),
                    (sourceModuleName, sourceCodeString.Code, ComponentType.StandardModule));
            }
            var actualModuleCode =  RefactoredCode(
                sourceModuleName,
                sourceCodeString.CaretPosition.ToOneBased(),
                presenterAction,
                null,
                false,
                ("StdModule", procedureModuleReferencingCode, ComponentType.StandardModule),
                (sourceModuleName, sourceCodeString.Code, ComponentType.StandardModule));

            return actualModuleCode;
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void ClassModuleUDTFieldSelection_ExternalReferences_ClassModule()
        {
            var sourceModuleName = "SourceModule";
            var sourceClassName = "theClass";
            var sourceModuleCode =
$@"

Public th|is As TBar";


            string classModuleReferencingCode =
$@"Option Explicit

Private {sourceClassName} As {sourceModuleName}
Private Const foo As String = ""Foo""
Private Const bar As Long = 7

Private Sub Class_Initialize()
    Set {sourceClassName} = New {sourceModuleName}
End Sub

Public Sub Foo()
    {sourceClassName}.this.First = foo
End Sub

Public Sub Bar()
    {sourceClassName}.this.Second = bar
End Sub

Public Sub FooBar()
    With {sourceClassName}
        .this.First = foo
        .this.Second = bar
    End With
End Sub
";

            var userInput = new UserInputDataObject()
                .UserSelectsField("this", "MyType");

            var presenterAction = Support.SetParameters(userInput);

            var sourceCodeString = sourceModuleCode.ToCodeString();

            var actualModuleCode = RefactoredCode(
                sourceModuleName,
                sourceCodeString.CaretPosition.ToOneBased(),
                presenterAction,
                null,
                false,
                ("ClassModule", classModuleReferencingCode, ComponentType.ClassModule),
                (sourceModuleName, sourceCodeString.Code, ComponentType.ClassModule));

            var referencingClassCode = actualModuleCode["ClassModule"];
            StringAssert.Contains($"{sourceClassName}.MyType.First = ", referencingClassCode);
            StringAssert.Contains($"{sourceClassName}.MyType.Second = ", referencingClassCode);
            StringAssert.Contains($"  .MyType.Second = ", referencingClassCode);
        }


        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void ClassModuleUDTFieldSelection_ExternalReferences_StdModule()
        {
            var sourceModuleName = "SourceModule";
            var sourceClassName = "theClass";
            var sourceModuleCode =
$@"

Public th|is As TBar";

            var procedureModuleReferencingCode =
$@"Option Explicit

Public Type TBar
    First As String
    Second As Long
End Type

Private {sourceClassName} As {sourceModuleName}
Private Const foo As String = ""Foo""
Private Const bar As Long = 7

Public Sub Initialize()
    Set {sourceClassName} = New {sourceModuleName}
End Sub

Public Sub Foo()
    {sourceClassName}.this.First = foo
End Sub

Public Sub Bar()
    {sourceClassName}.this.Second = bar
End Sub

Public Sub FooBar()
    With {sourceClassName}
        .this.First = foo
        .this.Second = bar
    End With
End Sub
";

            var userInput = new UserInputDataObject()
                .UserSelectsField("this", "MyType");

            var presenterAction = Support.SetParameters(userInput);

            var sourceCodeString = sourceModuleCode.ToCodeString();

            var actualModuleCode = RefactoredCode(
                sourceModuleName,
                sourceCodeString.CaretPosition.ToOneBased(),
                presenterAction,
                null,
                false,
                ("StdModule", procedureModuleReferencingCode, ComponentType.StandardModule),
                (sourceModuleName, sourceCodeString.Code, ComponentType.ClassModule));

            var referencingModuleCode = actualModuleCode["StdModule"];
            StringAssert.Contains($"{sourceClassName}.MyType.First = ", referencingModuleCode);
            StringAssert.Contains($"{sourceClassName}.MyType.Second = ", referencingModuleCode);
            StringAssert.Contains($"  .MyType.Second = ", referencingModuleCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UserDefinedTypeUserUpdatesToBeReadOnly()
        {
            string inputCode =
$@"
Private Type TBar
    First As String
    Second As Long
End Type

Private t|his As TBar";

            var userInput = new UserInputDataObject()
                .UserSelectsField("this", isReadOnly: true);

            var presenterAction = Support.SetParameters(userInput);

            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);

            StringAssert.DoesNotContain("Public Property Let First", actualCode);
            StringAssert.DoesNotContain("Public Property Let Second", actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void MultipleUserDefinedTypesOfSameName()
        {
            string inputCode =
$@"
Option Explicit

Private Type TBar
    FirstValue As Long
    SecondValue As String
End Type

Public mF|oo As TBar
";

            string module2Content =
$@"
Public Type TBar
    FirstVal As Long
    SecondVal As String
End Type
";

            var presenterAction = Support.UserAcceptsDefaults();

            var codeString = inputCode.ToCodeString();
            var actualModuleCode = RefactoredCode(
                "Module1",
                codeString.CaretPosition.ToOneBased(),
                presenterAction,
                null,
                false,
                ("Module2", module2Content, ComponentType.StandardModule),
                ("Module1", codeString.Code, ComponentType.StandardModule));

            var actualCode = actualModuleCode["Module1"];

            StringAssert.Contains($"Public Property Let FirstValue(", actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UDTMemberPropertyConflictsWithExistingFunction()
        {
            string inputCode =
$@"
Private Type TBar
    First As String
    Second As Long
End Type

Public myB|ar As TBar

Private Function First() As String
    First = myBar.First
End Function";

            var presenterAction = Support.UserAcceptsDefaults();

            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);

            StringAssert.Contains("Public Property Let First_1", actualCode);
            StringAssert.Contains("Public Property Let Second", actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UDTMemberIsPrivateUDT()
        {
            string inputCode =
$@"

Private Type TFoo
    Foo As Integer
    Bar As Byte
End Type

Private Type TBar
    FooBar As TFoo
End Type

Private my|Bar As TBar
";

            var presenterAction = Support.UserAcceptsDefaults();

            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);

            StringAssert.Contains("Public Property Let Foo(", actualCode);
            StringAssert.Contains("Public Property Let Bar(", actualCode);
            StringAssert.Contains($"myBar.FooBar.Foo = {Support.RHSIdentifier}", actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UDTMemberIsPrivateUDT_RepeatedType()
        {
            string inputCode =
$@"

Private Type TFoo
    Foo As Integer
    Bar As Byte
End Type

Private Type TBar
    FooBar As TFoo
    ReBar As TFoo
End Type

Private my|Bar As TBar
";

            var presenterAction = Support.UserAcceptsDefaults();

            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);

            StringAssert.Contains("Public Property Let Foo(", actualCode);
            StringAssert.Contains("Public Property Let Bar(", actualCode);
            StringAssert.Contains("Public Property Let Foo_1(", actualCode);
            StringAssert.Contains("Public Property Let Bar_1(", actualCode);
            StringAssert.Contains($"myBar.FooBar.Foo = {Support.RHSIdentifier}", actualCode);
            StringAssert.Contains($"myBar.ReBar.Foo = {Support.RHSIdentifier}", actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UDTMemberIsPublicUDT()
        {
            string inputCode =
$@"

Public Type TFoo
    Foo As Integer
    Bar As Byte
End Type

Private Type TBar
    FooBar As TFoo
End Type

Private my|Bar As TBar
";

            var presenterAction = Support.UserAcceptsDefaults();

            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);

            StringAssert.Contains("Public Property Let FooBar(", actualCode);
            StringAssert.Contains($"myBar.FooBar = {Support.RHSIdentifier}", actualCode);
        }

        protected override IRefactoring TestRefactoring(
            IRewritingManager rewritingManager,
            RubberduckParserState state,
            RefactoringUserInteraction<IEncapsulateFieldPresenter, EncapsulateFieldModel> userInteraction,
            ISelectionService selectionService)
        {
            return Support.SupportTestRefactoring(rewritingManager, state, userInteraction, selectionService);
        }
    }
}
