using NUnit.Framework;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.ImplicitTypeToExplicit;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;
using System;
using System.Collections.Generic;

namespace RubberduckTests.Refactoring.ImplicitTypeToExplicit
{
    [TestFixture]
    public class ImplicitTypeToExplicitRefactoringActionVariableTests : ImplicitTypeToExplicitRefactoringActionTestsBase
    {
        [TestCase("var1 = 42", "Long")]
        [TestCase("var1 = 42.5", "Double")]
        [TestCase("var1 = False", "Boolean")]
        [TestCase(@"var1 = ""StringLiteral""", "String")]
        [TestCase("var1 = #2015-05-15#", "Date")]
        [TestCase("var1 = 45E10", "Double")]
        [Category("Refactorings")]
        [Category(nameof(ImplicitTypeToExplicitRefactoringAction))]
        public void LocalVariable_AssignedUsingLiteralExpression(string assignment, string expectedType)
        {
            var targetName = "var1";
            var inputCode =
$@"Sub Foo()
    Dim var1
    {assignment}
End Sub";

            var refactoredCode = RefactoredCode(inputCode,
                state => TestModel(state, NameAndDeclarationTypeTuple(targetName), (model) => model));

            StringAssert.Contains($"{targetName} As {expectedType}", refactoredCode);
        }

        [TestCase("var1 = 42")]
        [TestCase("var1 = 42.5")]
        [Category("Refactorings")]
        [Category(nameof(ImplicitTypeToExplicitRefactoringAction))]
        public void RespectsForceVariantFlag(string assignment)
        {
            var targetName = "var1";
            var inputCode =
$@"Sub Foo()
    Dim var1
    {assignment}
End Sub";

            var refactoredCode = RefactoredCode(inputCode,
                state => TestModel(state, NameAndDeclarationTypeTuple(targetName), (model) => { model.ForceVariantAsType = true; return model; }));

            StringAssert.Contains($"{targetName} As Variant", refactoredCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(ImplicitTypeToExplicitRefactoringAction))]
        public void InvalidTarget_ThrowsArgumentException()
        {
            var targetName = "Foo";
            var inputCode =
$@"Sub Foo()
End Sub";

            Assert.Throws<ArgumentException>( () => RefactoredCode(inputCode,
                state => TestModel(state, (targetName, DeclarationType.Procedure), (model) => model)));
        }

        [TestCase("var1 = 42", "Long", "Long")]
        [TestCase("var1 = 42.55", "Long", "Double")]
        [Category("Refactorings")]
        [Category(nameof(ImplicitTypeToExplicitRefactoringAction))]
        public void LocalVariable_AssignedUsingLiteralExpresssionAndFunctions(string assignment, string functionType, string expectedType)
        {
            var targetName = "var1";
            var inputCode =
$@"Sub Foo()
    Dim var1
    {assignment}

    var1 = AssignAValue()
End Sub

Private Function AssignAValue() As {functionType}
End Function";

            var refactoredCode = RefactoredCode(inputCode,
                state => TestModel(state, NameAndDeclarationTypeTuple(targetName), (model) => model));

            StringAssert.Contains($"{targetName} As {expectedType}", refactoredCode);
        }

        [TestCase("mTest2", "String")]
        [TestCase("mTest", "Long")]
        [Category("Refactorings")]
        [Category(nameof(ImplicitTypeToExplicitRefactoringAction))]
        public void FixMutipleResults(string targetName, string expectedType)
        {
            var inputCode =
$@"
Private mTest2
Private mTest

Sub Foo(ByVal arg As String)
    mTest = AssignAValue()
    mTest2 = arg
End Sub

Private Function AssignAValue() As Long
End Function
";

            var refactoredCode = RefactoredCode(inputCode,
                state => TestModel(state, NameAndDeclarationTypeTuple(targetName), (model) => model));

            StringAssert.Contains($"{targetName} As {expectedType}", refactoredCode);
        }

        [Category("Refactorings")]
        [Category(nameof(ImplicitTypeToExplicitRefactoringAction))]
        public void SetAssignment_FromFunction()
        {
            var targetName = "mTest";
            var expectedType = "Collection";
            var inputCode =
$@"
Private mTest

Private mColl As Collection

Sub Fizz()
    Set mTest = AssignACollection()
End Sub

Private Function AssignACollection() As Collection
    Set AssignACollection = mColl
End Function";

            var refactoredCode = RefactoredCode(inputCode,
                state => TestModel(state, NameAndDeclarationTypeTuple(targetName), (model) => model));

            StringAssert.Contains($"{targetName} As {expectedType}", refactoredCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(ImplicitTypeToExplicitRefactoringAction))]
        public void SetAssignment_FromNew()
        {
            var targetName = "mTest";
            var expectedType = "Collection";
            var inputCode =
$@"
Private mTest

Sub Fizz()
    Set mTest = New Collection
End Sub
";

            var refactoredCode = RefactoredCode(inputCode,
                state => TestModel(state, NameAndDeclarationTypeTuple(targetName), (model) => model));

            StringAssert.Contains($"{targetName} As {expectedType}", refactoredCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(ImplicitTypeToExplicitRefactoringAction))]
        public void SetAssignment_FromOtherInstance()
        {
            var targetName = "mTest";
            var expectedType = "Collection";
            var inputCode =
$@"
Private mTest

Private mColl As Collection

Sub Fizz()
    Set mColl = New Collection
    Set mTest = mColl
End Sub
";

            var refactoredCode = RefactoredCode(inputCode,
                state => TestModel(state, NameAndDeclarationTypeTuple(targetName), (model) => model));

            StringAssert.Contains($"{targetName} As {expectedType}", refactoredCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(ImplicitTypeToExplicitRefactoringAction))]
        public void UseParameterType_Function()
        {
            var targetName = "mTest";
            var expectedType = "String";
            var inputCode =
$@"
Private mTest

Sub Fizz()
    Dim local As String
    local = AssignAValue(mTest)
End Sub

Private Function AssignAValue(arg As String) As String
    AssignAValue = arg & ""MoreContent""
End Function";

            var refactoredCode = RefactoredCode(inputCode,
                state => TestModel(state, NameAndDeclarationTypeTuple(targetName), (model) => model));

            StringAssert.Contains($"{targetName} As {expectedType}", refactoredCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(ImplicitTypeToExplicitRefactoringAction))]
        public void UseParameterType_EnumMember()
        {
            var targetName = "local";
            var expectedType = "Long";
            var inputCode =
$@"

Private Enum TestEnum
    FirstValue
    SecondValue
End Enum

Sub Fizz()
    Dim local
    local = FirstValue
End Sub
";

            var refactoredCode = RefactoredCode(inputCode,
                state => TestModel(state, NameAndDeclarationTypeTuple(targetName), (model) => model));

            StringAssert.Contains($"{targetName} As {expectedType}", refactoredCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(ImplicitTypeToExplicitRefactoringAction))]
        public void UseParameterType_Sub()
        {
            var targetName = "mTest";
            var expectedType = "String";
            var inputCode =
$@"
Private mTest

Sub Fizz()
    Dim local As String
    AssignAValue(mTest)
End Sub

Private Sub AssignAValue(ByRef arg As String)
    arg = arg & ""MoreContent""
End Sub";

            var refactoredCode = RefactoredCode(inputCode,
               state => TestModel(state, NameAndDeclarationTypeTuple(targetName), (model) => model));

            StringAssert.Contains($"{targetName} As {expectedType}", refactoredCode);
        }

        [TestCase("6, 7, mTest, 9")]
        [TestCase("arg:=mTest, bogey3:=9, bogey2:=7, bogey1:=6")]
        [Category("Refactorings")]
        [Category(nameof(ImplicitTypeToExplicitRefactoringAction))]
        public void UseParameterType_MultipleParameters(string argList)
        {
            var targetName = "mTest";
            var expectedType = "Double";
            var inputCode =
$@"
Private mTest

Sub Fizz()
    Dim local As String
    local = AssignAValue(mTest)
    local = AssignAValue2({argList})
End Sub

Private Function AssignAValue(arg As Integer) As String
    AssignAValue = CStr(arg)
End Function

Private Function AssignAValue2(bogey1 As Long, bogey2 As Long, arg As Double, bogey3 As Long) As String
    AssignAValue2 = CStr(arg)
End Function
";

            var refactoredCode = RefactoredCode(inputCode,
                state => TestModel(state, NameAndDeclarationTypeTuple(targetName), (model) => model));

            StringAssert.Contains($"{targetName} As {expectedType}", refactoredCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(ImplicitTypeToExplicitRefactoringAction))]
        public void UseParameterType_PartOfAParamArray()
        {
            var targetName = "mTest";
            var expectedType = "Double";
            var inputCode =
$@"
Private mTest

Sub Fizz()
    Dim local As String
    Dim anotherLocal As Double
    local = AssignAValue(6, anotherLocal, mTest)
End Sub

Private Function AssignAValue(arg As Integer, ByVal ParamArray args() As Double) As String
    AssignAValue = CStr(arg)
End Function
";

            var refactoredCode = RefactoredCode(inputCode,
                state => TestModel(state, NameAndDeclarationTypeTuple(targetName), (model) => model));

            StringAssert.Contains($"{targetName} As {expectedType}", refactoredCode);
        }

        [TestCase("String", "Long", "Variant")]
        [TestCase("String", "String", "String")]
        [TestCase("Date", "Long", "Variant")]
        [TestCase("Date", "Date", "Date")]
        [TestCase("Currency", "Long", "Variant")]
        [TestCase("Currency", "Currency", "Currency")]
        [Category("Refactorings")]
        [Category(nameof(ImplicitTypeToExplicitRefactoringAction))]
        public void ModuleVariable_RestrictiveAsTypes(string argType, string functionType, string expectedType)
        {
            var targetName = "mTest";
            var inputCode =
$@"
Private mTest

Sub Fizz()
    mTest = AssignAValue()
End Sub

Sub Fazz(arg As {argType})
    mTest = arg
End Sub

Private Function AssignAValue() As {functionType}
End Function";

            var refactoredCode = RefactoredCode(inputCode,
                state => TestModel(state, NameAndDeclarationTypeTuple(targetName), (model) => model));

            StringAssert.Contains($"{targetName} As {expectedType}", refactoredCode);
        }

        [TestCase("Public")]
        [TestCase("Private")]
        [Category("Refactorings")]
        [Category(nameof(ImplicitTypeToExplicitRefactoringAction))]
        public void ModuleVariable_AssignedUsingExternalField(string externalAccessibility)
        {
            var targetName = "mTest";
            var expectedType = "Long";
            var inputCode =
$@"
Public mTest
";
            var assigningModuleCode =
$@"
Private mFizz As Long
{externalAccessibility} Sub Fizz()
    mTest = mFizz
End Sub
";
            var vbe = MockVbeBuilder.BuildFromStdModules((MockVbeBuilder.TestModuleName, inputCode),
                ("AssigningModule", assigningModuleCode));
            var refactoredCode = RefactoredCode(vbe.Object,
                state => TestModel(state, NameAndDeclarationTypeTuple(targetName), (model) => model));

            StringAssert.Contains($"{targetName} As {expectedType}", refactoredCode[MockVbeBuilder.TestModuleName]);
        }

        [TestCase("Public")]
        [TestCase("Private")]
        [Category("Refactorings")]
        [Category(nameof(ImplicitTypeToExplicitRefactoringAction))]
        public void ModuleVariable_AssignedUsingExternalVariable(string externalAccessibility)
        {
            var targetName = "mTest";
            var expectedType = "Long";
            var inputCode =
$@"
Public mTest
";
            var assigningModuleCode =
$@"
{externalAccessibility} Sub Fizz()
    Dim localVar As Long
    localVar = 6
    mTest = localVar
End Sub
";
            var vbe = MockVbeBuilder.BuildFromStdModules((MockVbeBuilder.TestModuleName, inputCode),
                ("AssigningModule", assigningModuleCode));
            var refactoredCode = RefactoredCode(vbe.Object,
                state => TestModel(state, NameAndDeclarationTypeTuple(targetName), (model) => model));

            StringAssert.Contains($"{targetName} As {expectedType}", refactoredCode[MockVbeBuilder.TestModuleName]);
        }

        [TestCase("Public")]
        [TestCase("Private")]
        [Category("Refactorings")]
        [Category(nameof(ImplicitTypeToExplicitRefactoringAction))]
        public void ModuleVariable_AssignedUsingExternalFunction(string externalAccessibility)
        {
            var targetName = "mTest";
            var expectedType = "Long";
            var inputCode =
$@"
Public mTest
";
            var assigningModuleCode =
$@"
Private Sub Fizz()
    mTest = AValue
End Sub

{externalAccessibility} Property Get AValue() As Long
End Property
";
            var vbe = MockVbeBuilder.BuildFromStdModules((MockVbeBuilder.TestModuleName, inputCode),
                ("AssigningModule", assigningModuleCode));
            var refactoredCode = RefactoredCode(vbe.Object,
                state => TestModel(state, NameAndDeclarationTypeTuple(targetName), (model) => model));

            StringAssert.Contains($"{targetName} As {expectedType}", refactoredCode[MockVbeBuilder.TestModuleName]);
        }

        [TestCase("Public")]
        [TestCase("Private")]
        [Category("Refactorings")]
        [Category(nameof(ImplicitTypeToExplicitRefactoringAction))]
        public void ModuleVariable_AssignedUsingExternalConstant(string externalAccessibility)
        {
            var targetName = "mTest";
            var expectedType = "Long";
            var inputCode =
$@"
Public mTest
";
            var assigningModuleCode =
$@"
{externalAccessibility} Const ACONST As Long = 5
Private Sub Fizz()
    mTest = ACONST
End Sub
";
            var vbe = MockVbeBuilder.BuildFromStdModules((MockVbeBuilder.TestModuleName, inputCode),
                ("AssigningModule", assigningModuleCode));
            var refactoredCode = RefactoredCode(vbe.Object,
                state => TestModel(state, NameAndDeclarationTypeTuple(targetName), (model) => model));

            StringAssert.Contains($"{targetName} As {expectedType}", refactoredCode[MockVbeBuilder.TestModuleName]);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(ImplicitTypeToExplicitRefactoringAction))]
        public void Field_IsParameterOfExternalFunction()
        {
            var targetName = "mTest";
            var expectedType = "Long";
            var inputCode =
$@"
Private mTest

Private Sub Test()
    Fizz(mTest)
End Sub
";
            var assigningModuleCode =
$@"
Public Sub Fizz(arg As Long)
End Sub
";
            var vbe = MockVbeBuilder.BuildFromStdModules((MockVbeBuilder.TestModuleName, inputCode),
                ("AssigningModule", assigningModuleCode));
            var refactoredCode = RefactoredCode(vbe.Object,
                state => TestModel(state, NameAndDeclarationTypeTuple(targetName), (model) => model));

            StringAssert.Contains($"{targetName} As {expectedType}", refactoredCode[MockVbeBuilder.TestModuleName]);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(ImplicitTypeToExplicitRefactoringAction))]
        public void LocalVariable_IsParameterOfExternalFunction()
        {
            var targetName = "local";
            var expectedType = "Long";
            var inputCode =
$@"
Private Sub Test()
    Dim local
    Fizz(local)
End Sub
";
            var assigningModuleCode =
$@"
Public Sub Fizz(arg As Long)
End Sub
";
            var vbe = MockVbeBuilder.BuildFromStdModules((MockVbeBuilder.TestModuleName, inputCode),
                ("AssigningModule", assigningModuleCode));
            var refactoredCode = RefactoredCode(vbe.Object,
                state => TestModel(state, NameAndDeclarationTypeTuple(targetName), (model) => model));

            StringAssert.Contains($"{targetName} As {expectedType}", refactoredCode[MockVbeBuilder.TestModuleName]);
        }

        [TestCase("CInt", "Integer")]
        [TestCase("CLng", "Long")]
        [TestCase("CVar", "Variant")]
        [TestCase("CCur", "Currency")]
        [Category("Refactorings")]
        [Category(nameof(ImplicitTypeToExplicitRefactoringAction))]
        public void LocalVariable_ConversionFunctions(string vbaFunc, string expectedType)
        {
            var targetName = "local";
            var inputCode =
$@"
Private Sub Test(arg As String)
    Dim local
    local = {vbaFunc}(arg)
End Sub
";
            var vbe = MockVbeBuilder.BuildFromModules((MockVbeBuilder.TestModuleName, inputCode, ComponentType.StandardModule), ReferenceLibrary.VBA);
            var refactoredCode = RefactoredCode(vbe.Object,
                state => TestModel(state, NameAndDeclarationTypeTuple(targetName), (model) => model));

            StringAssert.Contains($"{targetName} As {expectedType}", refactoredCode[MockVbeBuilder.TestModuleName]);
        }

        [TestCase("CInt", "Integer")]
        [TestCase("CLng", "Long")]
        [TestCase("CVar", "Variant")]
        [TestCase("CCur", "Currency")]
        [Category("Refactorings")]
        [Category(nameof(ImplicitTypeToExplicitRefactoringAction))]
        public void Field_ConversionFunctions_OtherModule(string vbaFunc, string expectedType)
        {
            var targetName = "mTest";
            var inputCode =
$@"
Public mTest
";
            var assigningModuleName = "AssigningModule";
            var assigningModuleCode =
$@"
Private Sub Test(arg As String)
    mTest = {vbaFunc}(arg)
End Sub
";
            var modules = new List<(string name, string content, ComponentType componentType)>()
            {
                (MockVbeBuilder.TestModuleName, inputCode, ComponentType.StandardModule),
                (assigningModuleName, assigningModuleCode, ComponentType.StandardModule)
            };

            var vbe = MockVbeBuilder.BuildFromModules(modules, new ReferenceLibrary[] { ReferenceLibrary.VBA });
            var refactoredCode = RefactoredCode(vbe.Object,
                state => TestModel(state, NameAndDeclarationTypeTuple(targetName), (model) => model));

            StringAssert.Contains($"{targetName} As {expectedType}", refactoredCode[MockVbeBuilder.TestModuleName]);
        }

        //TODO: Once an expression evaluation capability is added, this should resolve to Double
        [Test]
        [Category("Refactorings")]
        [Category(nameof(ImplicitTypeToExplicitRefactoringAction))]
        public void AssignedByBooleanExpressionAndTypedField_Variant()
        {
            var targetName = "mTest";
            var expectedType = "Variant";
            var inputCode =
$@"
Private mTest
Private mTest2 As Double

Sub Test()
    mTest = 3 > 2
End Sub

Sub Test2()
    mTest = mTest2
End Sub
";

            var refactoredCode = RefactoredCode(inputCode,
                state => TestModel(state, NameAndDeclarationTypeTuple(targetName), (model) => model));

            StringAssert.Contains($"{targetName} As {expectedType}", refactoredCode);
        }

        //TODO: Once an expression evaluation capability is added, this should resolve to Boolean
        [Test]
        [Category("Refactorings")]
        [Category(nameof(ImplicitTypeToExplicitRefactoringAction))]
        public void BooleanNotExpression_UsingFunction()
        {
            var targetName = "mTest";
            var expectedType = "Variant";
            var inputCode =
$@"
Private mTest

Sub Test()
    mTest = Not GetBoolean()
End Sub

Private Function GetBoolean() As Boolean
    GetBoolean = true
End Function";

            var refactoredCode = RefactoredCode(inputCode,
                state => TestModel(state, NameAndDeclarationTypeTuple(targetName), (model) => model));

            StringAssert.Contains($"{targetName} As {expectedType}", refactoredCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(ImplicitTypeToExplicitRefactoringAction))]
        public void VariableTypedAsPropertyLetValueParameter()
        {
            var targetName = "local";
            var expectedType = "Double";
            var inputCode =
$@"
Private mTest As Double

Sub Test()
    Dim local
    local = 55
    AValue = local
End Sub

Private Property Let AValue(RHS As Double)
    mTest = RHS
End Property";

            var refactoredCode = RefactoredCode(inputCode,
                state => TestModel(state, NameAndDeclarationTypeTuple(targetName), (model) => model));

            StringAssert.Contains($"{targetName} As {expectedType}", refactoredCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(ImplicitTypeToExplicitRefactoringAction))]
        public void VariableTypedAsPropertyLetIndexParameter()
        {
            var targetName = "local";
            var expectedType = "Integer";
            var inputCode =
$@"
Private mTest As Double

Sub Test()
    Dim local
    AValue(local) = 44.4
End Sub

Private Property Set AValue(index As Integer, RHS As Double)
    mTest = RHS
End Property";

            var refactoredCode = RefactoredCode(inputCode,
                state => TestModel(state, NameAndDeclarationTypeTuple(targetName), (model) => model));

            StringAssert.Contains($"{targetName} As {expectedType}", refactoredCode);
        }

        [TestCase("", "Boolean")]
        [TestCase("ModuleOne.", "Double")]
        [TestCase("ModuleTwo.", "String")]
        [TestCase("", "Boolean")]
        [Category("Refactorings")]
        [Category(nameof(ImplicitTypeToExplicitRefactoringAction))]
        public void VariableTypedAsPropertyLetValueParameterNameRespectsMemberAccessExpr(string moduleQualification, string expectedType)
        {
            var externalModule1Name = "ModuleOne";
            var externalModule1Code =
$@"
Public Property Let AValue(ByVal RHS As Double)
End Property
";

            var externalModule2Name = "ModuleTwo";
            var externalModule2Code =
$@"
Public Property Let AValue(ByVal RHS As String)
End Property
";

            var targetName = "this";
            var inputCode =
$@"
Private Sub Test()
    Dim this
    {moduleQualification}AValue = this
End Sub

Public Property Let AValue(ByVal RHS As Boolean)
End Property
";
            var vbe = MockVbeBuilder.BuildFromStdModules((MockVbeBuilder.TestModuleName, inputCode),
                (externalModule1Name, externalModule1Code),
                (externalModule2Name, externalModule2Code));

            var refactoredCode = RefactoredCode(vbe.Object,
                state => TestModel(state, NameAndDeclarationTypeTuple(targetName), (model) => model));

            StringAssert.Contains($"{targetName} As {expectedType}", refactoredCode[MockVbeBuilder.TestModuleName]);
        }

        [TestCase("", "Boolean")]
        [TestCase("With ModuleOne", "Double")]
        [TestCase("With ModuleTwo", "String")]
        [TestCase("", "Boolean")]
        [Category("Refactorings")]
        [Category(nameof(ImplicitTypeToExplicitRefactoringAction))]
        public void VariableTypedAsPropertyLetValueParameterNameRespectsWithAccessExpr(string moduleQualification, string expectedType)
        {
            var endWith = moduleQualification.Length > 0 ? "End With" : string.Empty;
            var dot = moduleQualification.Length > 0 ? "." : string.Empty;

            var externalModule1Name = "ModuleOne";
            var externalModule1Code =
$@"
Public Property Let AValue(ByVal RHS As Double)
End Property
";

            var externalModule2Name = "ModuleTwo";
            var externalModule2Code =
$@"
Public Property Let AValue(ByVal RHS As String)
End Property
";

            var targetName = "this";
            var inputCode =
$@"
Private Sub Test()
    Dim this
    {moduleQualification}
        {dot}AValue = this
    {endWith}
End Sub

Public Property Let AValue(ByVal RHS As Boolean)
End Property
";
            var vbe = MockVbeBuilder.BuildFromStdModules((MockVbeBuilder.TestModuleName, inputCode),
                (externalModule1Name, externalModule1Code),
                (externalModule2Name, externalModule2Code));

            var refactoredCode = RefactoredCode(vbe.Object,
                state => TestModel(state, NameAndDeclarationTypeTuple(targetName), (model) => model));

            StringAssert.Contains($"{targetName} As {expectedType}", refactoredCode[MockVbeBuilder.TestModuleName]);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(ImplicitTypeToExplicitRefactoringAction))]
        public void VariableTypedAsPropertyLetValueParameterResolveMultipleTypes()
        {
            var targetName = "this";
            var expectedType = "Variant";
            var inputCode =
$@"
Private Sub Test()
    Dim this
    AValue = this
    AValue1 = this
    AValue2 = this
End Sub

Public Property Let AValue(ByVal RHS As Boolean)
End Property

Public Property Let AValue1(ByVal RHS As String)
End Property

Public Property Let AValue2(ByVal RHS As Double)
End Property
";
            var vbe = MockVbeBuilder.BuildFromStdModules((MockVbeBuilder.TestModuleName, inputCode));

            var refactoredCode = RefactoredCode(vbe.Object,
                state => TestModel(state, NameAndDeclarationTypeTuple(targetName), (model) => model));

            StringAssert.Contains($"{targetName} As {expectedType}", refactoredCode[MockVbeBuilder.TestModuleName]);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(ImplicitTypeToExplicitRefactoringAction))]
        public void VariableTypedByLiteralConcedesToExplicitReferencedType()
        {
            var targetName = "this";
            var expectedType = "Byte";
            var inputCode =
$@"
Private Sub Test()
    Dim this
    this = 55
    AValue2 = this
End Sub

Public Property Let AValue2(ByVal RHS As Byte)
End Property
";
            var vbe = MockVbeBuilder.BuildFromStdModules((MockVbeBuilder.TestModuleName, inputCode));

            var refactoredCode = RefactoredCode(vbe.Object,
                state => TestModel(state, NameAndDeclarationTypeTuple(targetName), (model) => model));

            StringAssert.Contains($"{targetName} As {expectedType}", refactoredCode[MockVbeBuilder.TestModuleName]);
        }

        [TestCase("Property", "Get ")]
        [TestCase("Function", "")]
        [Category("Refactorings")]
        [Category(nameof(ImplicitTypeToExplicitRefactoringAction))]
        public void VariableTypedByFunctionType(string procedureType, string getToken)
        {
            var targetName = "mTest";
            var expectedType = "Long";
            var inputCode =
$@"
Public mTest

Public {procedureType} {getToken}AValue() As Long
    AValue = mTest
End {procedureType}
";
            var refactoredCode = RefactoredCode(inputCode, NameAndDeclarationTypeTuple(targetName));
            StringAssert.Contains($"{targetName} As {expectedType}", refactoredCode);
        }

        //https://github.com/rubberduck-vba/Rubberduck/issues/5646        
        [TestCase("()")]
        [TestCase("(1,2,3,4,5)")]
        [Category("Refactorings")]
        [Category(nameof(ImplicitTypeToExplicitRefactoringAction))]
        public void ArrayVariable(string arrayDimensionExpression)
        {
            var targetName = "mTest";
            var expectedType = "Variant";
            var inputCode = $"Public {targetName}{arrayDimensionExpression}";

            var refactoredCode = RefactoredCode(inputCode, NameAndDeclarationTypeTuple(targetName));
            StringAssert.Contains($"{targetName}{arrayDimensionExpression} As {expectedType}", refactoredCode);
        }

        //https://github.com/rubberduck-vba/Rubberduck/issues/5907
        [Test]
        [Category("Refactorings")]
        [Category(nameof(ImplicitTypeToExplicitRefactoringAction))]
        public void StringConcatOpWithLineContinuation()
        {
            var targetName = "testValue";
            var expectedType = "String";
            var inputCode =
$@"
Private Sub Test()
    Dim testValue
    testValue = ""Hello "" & _
        ""World""
End Sub
";
            var vbe = MockVbeBuilder.BuildFromStdModules((MockVbeBuilder.TestModuleName, inputCode));

            var refactoredCode = RefactoredCode(vbe.Object,
                state => TestModel(state, NameAndDeclarationTypeTuple(targetName), (model) => model));

            StringAssert.Contains($"{targetName} As {expectedType}", refactoredCode[MockVbeBuilder.TestModuleName]);
        }

        [TestCase(@"""Literal1 "" & ""Literal2""", "String")]
        [TestCase(@"""Literal1 "" & 5", "String")]
        [TestCase(@"""Literal1 "" & 5 & "" Literal2""", "String")]
        [TestCase(@"""Literal1 "" & declaredVariant & "" Literal2""", "Variant")]
        [TestCase(@"""Literal1 "" & GetAString() & "" Literal2""", "String")]
        [TestCase(@"GetAString() & ""Literal1 ""  & "" Literal2""", "String")]
        [TestCase(@"GetALong() & ""Literal1 ""  & "" Literal2""", "String")]
        [TestCase(@"doubleVal & ""Literal1 ""  & "" Literal2""", "String")]
        [TestCase(@"PI & ""Literal1 ""  & "" Literal2""", "String")]
        [TestCase(@"implicitVar & ""Literal1 ""  & "" Literal2""", "Variant")]
        [TestCase(@"PI & GetAString()  & GetALong() & doubleVal & ""Literal1""", "String")]
        [TestCase(@"PI & GetAString()  & GetALong() & doubleVal & implicitVar & ""Literal1""", "Variant")]
        [Category("Refactorings")]
        [Category(nameof(ImplicitTypeToExplicitRefactoringAction))]
        public void ConcatOpContext(string expression, string expectedType)
        {
            var targetName = "testValue";
            var inputCode =
$@"

Private Const PI As Single = 3.14

Private Sub Test()
    Dim testValue, implicitVar
    Dim declaredVariant As Variant
    declaredVariant = ""I'm a Variant""

    Dim doubleVal As Double

    testValue = {expression}
End Sub

Private Function GetAString() As String
    GetAString = ""Rubberduck""
End Function

Private Function GetALong() As Long
    GetALong = 2022
End Function
";
            var vbe = MockVbeBuilder.BuildFromStdModules((MockVbeBuilder.TestModuleName, inputCode));

            var refactoredCode = RefactoredCode(vbe.Object,
                state => TestModel(state, NameAndDeclarationTypeTuple(targetName), (model) => model));

            StringAssert.Contains($"{targetName} As {expectedType}", refactoredCode[MockVbeBuilder.TestModuleName]);
        }

        [TestCase(@"GetAValue() & ""Literal""", "GetAValue() As Long", "GetAValue() As Variant", "String")]
        [TestCase(@"GetAValue() & ""Literal""", "GetAValue() As Variant", "GetAValue() As Long", "Variant")]
        [Category("Refactorings")]
        [Category(nameof(ImplicitTypeToExplicitRefactoringAction))]
        public void ConcatOpContextUsesCorrectFunctionDeclaration(string expression, string sameNameFunction, 
            string otherModuleDeclaration, string expectedType)
        {
            var targetName = "testValue";
            var inputCode =
$@"
Private Sub Test()
    Dim testValue

    testValue = {expression}
End Sub

Private Function {sameNameFunction}
    GetAValue = 2022
End Function
";

            //Create another module with identical declaration identifiers but a different
            //AsTypeName for the function.  The refactoring targetName is changed to allow
            //NameAndDeclarationTypeTuple to succeed.
            var otherModuleCode = inputCode.Replace(sameNameFunction, otherModuleDeclaration);
            
            otherModuleCode = otherModuleCode.Replace(targetName, "testValue2");

            var vbe = MockVbeBuilder.BuildFromStdModules((MockVbeBuilder.TestModuleName, inputCode),
                ("OtherModule", otherModuleCode));

            var refactoredCode = RefactoredCode(vbe.Object,
                state => TestModel(state, NameAndDeclarationTypeTuple(targetName), (model) => model));

            StringAssert.Contains($"{targetName} As {expectedType}", refactoredCode[MockVbeBuilder.TestModuleName]);
        }

        [TestCase(@"PI & "" Literal""", "Const PI As Single = 3.14", "Const PI As Variant = 3.14", "String")]
        [TestCase(@"PI & "" Literal""", "Const PI As Variant = 3.14", "Const PI As Single = 3.14", "Variant")]
        [Category("Refactorings")]
        [Category(nameof(ImplicitTypeToExplicitRefactoringAction))]
        public void ConcatOpContextUsesCorrectConstantDeclaration(string expression, string sameNameConst, 
            string otherModuleConst, string expectedType)
        {
            var targetName = "testValue";
            var inputCode =
$@"
Private {sameNameConst}

Private Sub Test()
    Dim testValue

    testValue = {expression}
End Sub
";

            //Create another module with identical declaration identifiers but a different
            //AsTypeName for the constant.  The refactoring targetName is changed to allow
            //NameAndDeclarationTypeTuple to succeed.
            var otherModuleCode = inputCode.Replace(sameNameConst, otherModuleConst);
            
            otherModuleCode = otherModuleCode.Replace(targetName, "testValue2");

            var vbe = MockVbeBuilder.BuildFromStdModules((MockVbeBuilder.TestModuleName, inputCode),
                ("OtherModule", otherModuleCode));

            var refactoredCode = RefactoredCode(vbe.Object,
                state => TestModel(state, NameAndDeclarationTypeTuple(targetName), (model) => model));

            StringAssert.Contains($"{targetName} As {expectedType}", refactoredCode[MockVbeBuilder.TestModuleName]);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(ImplicitTypeToExplicitRefactoringAction))]
        public void ConcatExpressionEdgeCase()
        {
            var targetName = "testValue";
            //TODO: Once an expression evaluation capability is added, 
            //expectedType should be 'String'
            var expectedType = "Variant";
            var inputCode =
$@"
Private Sub Test()
    Dim testValue

    testValue = 5 & Null & Null & 5
End Sub
";

            var refactoredCode = RefactoredCode(inputCode,
                state => TestModel(state, NameAndDeclarationTypeTuple(targetName), (model) => model));

            StringAssert.Contains($"{targetName} As {expectedType}", refactoredCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(ImplicitTypeToExplicitRefactoringAction))]
        public void ConcatExpressionContainsNonLExprOrLiteralContext()
        {
            var targetName = "testValue";
            //TODO: Once an expression evaluation capability is added, 
            //expectedType should be 'String'
            var expectedType = "Variant";
            var inputCode =
$@"
Private Sub Test()
    Dim testValue
    testValue = ""Literal"" & 55 + 20
End Sub
";

            var refactoredCode = RefactoredCode(inputCode,
                state => TestModel(state, NameAndDeclarationTypeTuple(targetName), (model) => model));

            StringAssert.Contains($"{targetName} As {expectedType}", refactoredCode);
        }

        [TestCase("Variant", "Variant")]
        [TestCase("Long", "String")]
        [TestCase("Double", "String")]
        [Category("Refactorings")]
        [Category(nameof(ImplicitTypeToExplicitRefactoringAction))]
        public void MultipleConcatOpExpressions(string typeName, string expectedType)
        {
            var targetName = "mTestValue";
            var inputCode =
$@"
Private mTestValue

Private Sub Test1()
    mTestValue = 5 & Null
End Sub

Private Sub Test2()
    Dim xVar As {typeName}
    xVar = 5
    mTestValue = xVar & ""Test2""
End Sub

Private Sub Test3()
    Dim xVar As Long
    xVar = 5
    mTestValue = xVar & ""Test3""
End Sub
";

            var refactoredCode = RefactoredCode(inputCode,
                state => TestModel(state, NameAndDeclarationTypeTuple(targetName), (model) => model));

            StringAssert.Contains($"{targetName} As {expectedType}", refactoredCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(ImplicitTypeToExplicitRefactoringAction))]
        public void ObjectWithinConcatExpressions()
        {
            var targetName = "mTestValue";
            var expectedType = "Variant";
            var inputCode =
@"
Private mTestValue

Private Sub Test1()
    Dim obj As TestClass
    set obj = new TestClass
    mTestValue = 5 & obj
End Sub
";

            var testClassCode =
$@"
Public mVal As String
";

            var vbe = TestVbe(("TestModule", inputCode, ComponentType.StandardModule),
                ("TestClass", testClassCode, ComponentType.ClassModule));


            var refactoredCode = RefactoredCode(vbe,
                state => TestModel(state, NameAndDeclarationTypeTuple(targetName), (model) => model));

            StringAssert.Contains($"{targetName} As {expectedType}", refactoredCode["TestModule"]);
        }

        private static (string, DeclarationType) NameAndDeclarationTypeTuple(string name)
            => (name, DeclarationType.Variable);
    }
}
