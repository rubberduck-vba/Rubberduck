using NUnit.Framework;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.ImplicitTypeToExplicit;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;
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

        private static (string, DeclarationType) NameAndDeclarationTypeTuple(string name)
            => (name, DeclarationType.Variable);
    }
}
