using NUnit.Framework;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.Utility;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace RubberduckTests.Refactoring.EncapsulateField
{
    [TestFixture]
    public class EncapsulatedUDTField : InteractiveRefactoringTestBase<IEncapsulateFieldPresenter, EncapsulateFieldModel>
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
            StringAssert.Contains("Private this1 As TBar", actualCode);
            StringAssert.DoesNotContain("this1 = value", actualCode);
            StringAssert.DoesNotContain($"This = this1", actualCode);
            StringAssert.Contains($"Public Property Get First", actualCode);
            StringAssert.Contains($"Public Property Get Second", actualCode);
            StringAssert.Contains($"this1.First = value", actualCode);
            StringAssert.Contains($"First = this1.First", actualCode);
            StringAssert.Contains($"this1.Second = value", actualCode);
            StringAssert.Contains($"Second = this1.Second", actualCode);
        }


        [TestCase("Public")]
        [TestCase("Private")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UserDefinedType_UserAcceptsDefaults_TwoUDTInstances(string accessibility)
        {
            string selectionModule = "ClassInput";
            string inputCode =
$@"
Private Type TBar
    First As String
    Second As Long
End Type

{accessibility} th|is As TBar

Private that As TBar";

            string otherModule = "OtherModule";
            string moduleCode =
$@"
Private Type TBar
    First As String
    Second As Long
End Type

{accessibility} this As TBar

Private that As TBar";


            var presenterAdjustment = Support.UserAcceptsDefaults();
            var codeString = inputCode.ToCodeString();
            var actualCode = RefactoredCode(selectionModule, codeString.CaretPosition.ToOneBased(), presenterAdjustment, null, false, (selectionModule, codeString.Code, ComponentType.ClassModule), (otherModule, moduleCode, ComponentType.StandardModule));
            StringAssert.Contains("Private this1 As TBar", actualCode[selectionModule]);
            StringAssert.DoesNotContain("this1 = value", actualCode[selectionModule]);
            StringAssert.DoesNotContain($"This = this1", actualCode[selectionModule]);
            StringAssert.DoesNotContain($"Public Property Get This(", actualCode[selectionModule]);
            StringAssert.Contains($"Public Property Get This_First", actualCode[selectionModule]);
            StringAssert.Contains($"Public Property Get This_Second", actualCode[selectionModule]);
            StringAssert.Contains($"this1.First = value", actualCode[selectionModule]);
            StringAssert.Contains($"This_First = this1.First", actualCode[selectionModule]);
            StringAssert.Contains($"this1.Second = value", actualCode[selectionModule]);
            StringAssert.Contains($"This_Second = this1.Second", actualCode[selectionModule]);
        }


        [TestCase("Public", true, true)]
        [TestCase("Private", true, true)]
        [TestCase("Public", false, false)]
        [TestCase("Private", false, false)]
        [TestCase("Public", true, false)]
        [TestCase("Private", true, false)]
        [TestCase("Public", false, true)]
        [TestCase("Private", false, true)]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UserDefinedType_TwoInstanceVariables(string accessibility, bool encapsulateThis, bool encapsulateThat)
        {
            string inputCode =
$@"
Private Type TBar
    First As String
    Second As Long
End Type

{accessibility} th|is As TBar
{accessibility} that As TBar";

            var userInput = new UserInputDataObject("this", "MyType", encapsulateThis);
            userInput.AddAttributeSet("that", "MyOtherType", encapsulateThat);

            var expectedThis = new EncapsulationIdentifiers("this") { Property = "MyType" };
            var expectedThat = new EncapsulationIdentifiers("that") { Property = "MyOtherType" };

            var presenterAction = Support.SetParameters(userInput);
            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            if (encapsulateThis)
            {
                StringAssert.Contains($"Private {expectedThis.Field} As TBar", actualCode);
                StringAssert.Contains($"This_First = {expectedThis.Field}.First", actualCode);
                StringAssert.Contains($"This_Second = {expectedThis.Field}.Second", actualCode);
                StringAssert.Contains($"{expectedThis.Field}.First = value", actualCode);
                StringAssert.Contains($"{expectedThis.Field}.Second = value", actualCode);
                StringAssert.Contains($"Property Get This_First", actualCode);
                StringAssert.Contains($"Property Get This_Second", actualCode);
            }
            else
            {
                StringAssert.Contains($"{accessibility} this As TBar", actualCode);
            }

            if (encapsulateThat)
            {
                StringAssert.Contains($"Private {expectedThat.Field} As TBar", actualCode);
                StringAssert.Contains($"That_First = {expectedThat.Field}.First", actualCode);
                StringAssert.Contains($"That_Second = {expectedThat.Field}.Second", actualCode);
                StringAssert.Contains($"{expectedThat.Field}.First = value", actualCode);
                StringAssert.Contains($"{expectedThat.Field}.Second = value", actualCode);
                StringAssert.Contains($"Property Get That_First", actualCode);
                StringAssert.Contains($"Property Get That_Second", actualCode);
            }
            else
            {
                StringAssert.Contains($"{accessibility} that As TBar", actualCode);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UserDefinedTypeMembers_OnlyEncapsulateUDTMembers()
        {
            string inputCode =
$@"
Private Type TBar
    First As String
    Second As Long
End Type

Private th|is As TBar";


            var userInput = new UserInputDataObject("this");

            var presenterAction = Support.SetParameters(userInput);

            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            StringAssert.Contains("this1.First = value", actualCode);
            StringAssert.Contains($"First = this1.First", actualCode);
            StringAssert.Contains("this1.Second = value", actualCode);
            StringAssert.Contains($"Second = this1.Second", actualCode);
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


            var userInput = new UserInputDataObject("this", "MyType", true);
            userInput.AddAttributeSet("mFoo", "Foo", true);
            userInput.AddAttributeSet("mBar", "Bar", true);
            userInput.AddAttributeSet("mFizz", "Fizz", true);

            var presenterAction = Support.SetParameters(userInput);

            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);

            StringAssert.Contains("Private this As TBar", actualCode);
            StringAssert.DoesNotContain("this = value", actualCode);
            StringAssert.DoesNotContain("MyType = this", actualCode);
            StringAssert.Contains($"this.First = value", actualCode);
            StringAssert.Contains($"First = this.First", actualCode);
            StringAssert.Contains($"this.Second = value", actualCode);
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

            var userInput = new UserInputDataObject("this", "MyType", true);

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
            StringAssert.DoesNotContain("this = value", actualCode);
            StringAssert.DoesNotContain("MyType = this", actualCode);
            StringAssert.Contains("Property Set First(ByVal value As Class1)", actualCode);
            StringAssert.Contains("Property Get First() As Class1", actualCode);
            StringAssert.Contains($"Set this.First = value", actualCode);
            StringAssert.Contains($"Set First = this.First", actualCode);
            StringAssert.Contains($"this.Second = value", actualCode);
            StringAssert.Contains($"Second = this.Second", actualCode);
            StringAssert.DoesNotContain($"Second = Second", actualCode);
            Assert.AreEqual(actualCode.IndexOf("this.First = value"), actualCode.LastIndexOf("this.First = value"));
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

            var userInput = new UserInputDataObject("this", "MyType");

            var presenterAction = Support.SetParameters(userInput);

            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            StringAssert.Contains("Private this As TBar", actualCode);
            StringAssert.DoesNotContain("this = value", actualCode);
            StringAssert.DoesNotContain("MyType = this", actualCode);
            StringAssert.Contains($"this.First = value", actualCode);
            StringAssert.Contains($"First = this.First", actualCode);
            StringAssert.Contains($"IsObject", actualCode);
            StringAssert.Contains($"this.Second = value", actualCode);
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

            var userInput = new UserInputDataObject("this", "MyType");

            var presenterAction = Support.SetParameters(userInput);

            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            StringAssert.Contains($"Private this As TBar", actualCode);
            StringAssert.DoesNotContain("this = value", actualCode);
            StringAssert.DoesNotContain($"this.First = value", actualCode);
            StringAssert.Contains($"Property Get First() As Variant", actualCode);
            StringAssert.Contains($"First = this.First", actualCode);
            StringAssert.DoesNotContain($"this.Second = value", actualCode);
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

            var userInput = new UserInputDataObject("this", "MyType", true);

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
            {
                StringAssert.Contains("this = value", actualCode);
                StringAssert.Contains("MyType = this", actualCode);
                StringAssert.Contains($"Public Property Get MyType", actualCode);
                StringAssert.Contains($"Private this As TBar", actualCode);
                Assert.AreEqual(actualCode.IndexOf("MyType = this"), actualCode.LastIndexOf("MyType = this"));
                StringAssert.DoesNotContain($"this.First = value", actualCode);
                StringAssert.DoesNotContain($"this.Second = value", actualCode);
                StringAssert.DoesNotContain($"Second = Second", actualCode);
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

            var defaults = new EncapsulationIdentifiers("mTheClass");

            StringAssert.Contains($"Private {defaults.Field} As Class1", actualCode);
            StringAssert.Contains($"Set {defaults.Field} = value", actualCode);
            StringAssert.Contains($"MTheClass = {defaults.Field}", actualCode);
            StringAssert.Contains($"Public Property Set {defaults.Property}", actualCode);
        }

        [TestCase(false)]
        [TestCase(true)]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void StandardModuleSource_UDTFieldSelection_ExternalReferences(bool moduleQualifyReference)
        {
            var userInput = new UserInputDataObject("this", "MyType");
            var sourceModuleName = "SourceModule";

            var actualModuleCode = StdModuleSource_TwoReferencingModulesScenario(moduleQualifyReference, sourceModuleName, userInput);

            var referencingModuleCode = actualModuleCode["StdModule"];
            StringAssert.Contains($"{sourceModuleName}.MyType.First = ", referencingModuleCode);
            StringAssert.Contains($"{sourceModuleName}.MyType.Second = ", referencingModuleCode);
            StringAssert.Contains($"  .MyType.Second = ", referencingModuleCode);

            var referencingClassCode = actualModuleCode["ClassModule"];
            StringAssert.Contains($"{sourceModuleName}.MyType.First = ", referencingClassCode);
            StringAssert.Contains($"{sourceModuleName}.MyType.Second = ", referencingClassCode);
            StringAssert.Contains($"  .MyType.Second = ", referencingClassCode);
        }

        private IDictionary<string,string> StdModuleSource_TwoReferencingModulesScenario(bool moduleResolve, string sourceModuleName, UserInputDataObject userInput)
        {
            var referenceExpression = moduleResolve ? $"{sourceModuleName}.this" : "this";
            var sourceModuleCode =
$@"
Public Type TBar
    First As String
    Second As Long
End Type

Public th|is As TBar";

            var procedureModuleReferencingCode =
$@"Option Explicit

Private Const foo As String = ""Foo""

Private Const bar As Long = 7

Public Sub Foo()
    {referenceExpression}.First = foo
End Sub

Public Sub Bar()
    {referenceExpression}.Second = bar
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

Private Const foo As String = ""Foo""

Private Const bar As Long = 7

Public Sub Foo()
    {referenceExpression}.First = foo
End Sub

Public Sub Bar()
    {referenceExpression}.Second = bar
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
            var actualModuleCode =  RefactoredCode(
                sourceModuleName,
                sourceCodeString.CaretPosition.ToOneBased(),
                presenterAction,
                null,
                false,
                ("StdModule", procedureModuleReferencingCode, ComponentType.StandardModule),
                ("ClassModule", classModuleReferencingCode, ComponentType.ClassModule),
                (sourceModuleName, sourceCodeString.Code, ComponentType.StandardModule));

            return actualModuleCode;
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void ClassModuleSource_UDTFieldSelection_ExternalReferences()
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

            var userInput = new UserInputDataObject("this", "MyType");

            var presenterAction = Support.SetParameters(userInput);

            var sourceCodeString = sourceModuleCode.ToCodeString();

            var actualModuleCode = RefactoredCode(
                sourceModuleName,
                sourceCodeString.CaretPosition.ToOneBased(),
                presenterAction,
                null,
                false,
                ("StdModule", procedureModuleReferencingCode, ComponentType.StandardModule),
                ("ClassModule", classModuleReferencingCode, ComponentType.ClassModule),
                (sourceModuleName, sourceCodeString.Code, ComponentType.ClassModule));

            var referencingModuleCode = actualModuleCode["StdModule"];
            StringAssert.Contains($"{sourceClassName}.MyType.First = ", referencingModuleCode);
            StringAssert.Contains($"{sourceClassName}.MyType.Second = ", referencingModuleCode);
            StringAssert.Contains($"  .MyType.Second = ", referencingModuleCode);

            var referencingClassCode = actualModuleCode["ClassModule"];
            StringAssert.Contains($"{sourceClassName}.MyType.First = ", referencingClassCode);
            StringAssert.Contains($"{sourceClassName}.MyType.Second = ", referencingClassCode);
            StringAssert.Contains($"  .MyType.Second = ", referencingClassCode);
        }




        protected override IRefactoring TestRefactoring(IRewritingManager rewritingManager, RubberduckParserState state, IRefactoringPresenterFactory factory, ISelectionService selectionService)
        {
            return Support.SupportTestRefactoring(rewritingManager, state, factory, selectionService);
        }
    }
}
