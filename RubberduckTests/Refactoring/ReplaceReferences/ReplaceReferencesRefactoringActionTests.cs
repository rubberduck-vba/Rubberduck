using NUnit.Framework;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.ReplaceReferences;
using Rubberduck.Refactorings.ReplacePrivateUDTMemberReferences;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;
using System.Linq;

namespace RubberduckTests.Refactoring.RenameReferences
{
    [TestFixture]
    public class ReplaceReferencesRefactoringActionTests : RefactoringActionTestBase<ReplaceReferencesModel>
    {
        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(ReplaceReferencesRefactoringAction))]
        public void ValueField()
        {
            var inputCode =
$@"
Option Explicit

Public mTest As Long

Private mValue As Long

Public Sub Fizz(arg As Long)
    mValue = mTest + arg
End Sub

Public Function RetrieveTest() As Long
    RetrieveTest = mTest
End Function
";

            string externalModule = "ExternalModule";
            string externalCode =
$@"
Option Explicit

Private mValue As Long

Public Sub Fazz(arg As Long)
    mValue = mTest * arg
End Sub
";
            var vbe = MockVbeBuilder.BuildFromStdModules((MockVbeBuilder.TestModuleName, inputCode), (externalModule, externalCode));
            var results = RefactoredCode(vbe.Object, state => TestModel(state, ("mTest", "mNewName", "RetrieveTest()")));

            StringAssert.Contains($"RetrieveTest = mNewName", results[MockVbeBuilder.TestModuleName]);
            StringAssert.Contains($"Public mTest As Long", results[MockVbeBuilder.TestModuleName]);

            StringAssert.Contains($"mValue = {MockVbeBuilder.TestModuleName}.RetrieveTest()", results[externalModule]);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(ReplaceReferencesRefactoringAction))]
        public void ExternalReferences()
        {
            var sourceModuleName = "SourceModule";
            var referenceExpression = $"{sourceModuleName}.";
            var sourceModuleCode =
$@"

Public this As Long";

            var procedureModuleReferencingCode =
$@"Option Explicit

Private Const bar As Long = 7

Public Sub Bar()
    {referenceExpression}this = bar
End Sub

Public Sub Foo()
    With {sourceModuleName}
        .this = bar
    End With
End Sub
";

            var vbe = MockVbeBuilder.BuildFromStdModules((sourceModuleName, sourceModuleCode), (MockVbeBuilder.TestModuleName, procedureModuleReferencingCode));
            var actualModuleCode = RefactoredCode(vbe.Object, state => TestModel(state, ("this", "test", "MyProperty")));

            var referencingModuleCode = actualModuleCode[MockVbeBuilder.TestModuleName];
            StringAssert.Contains($"{sourceModuleName}.MyProperty = ", referencingModuleCode);
            StringAssert.DoesNotContain($"{sourceModuleName}.{sourceModuleName}.MyProperty = ", referencingModuleCode);
            StringAssert.Contains($"  .MyProperty = bar", referencingModuleCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(ReplaceReferencesRefactoringAction))]
        public void ArrayReferences2()
        {
            var sourceModuleName = "SourceModule";
            var inputCode =
    @"Private Sub Foo()
    ReDim arr(0 To 1)
    arr(1) = arr(0)
End Sub";
            string expectedCode =
                    @"Private Sub Foo()
    ReDim bar(0 To 1)
    bar(1) = bar(0)
End Sub";

            var vbe = MockVbeBuilder.BuildFromStdModules((sourceModuleName, inputCode));
            var actualModuleCode = RefactoredCode(vbe.Object, state => TestModel(state, ("arr", "bar", null)));

            var results = actualModuleCode[sourceModuleName];
            Assert.AreEqual(expectedCode, results);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(ReplaceReferencesRefactoringAction))]
        public void ArrayReferences()
        {
            var sourceModuleName = "SourceModule";
            var inputCode =
$@"
Option Explicit

Private myArray() As Integer

Private Sub InitializeArray(size As Long)
    Redim myArray(size)
    Dim idx As Long
    For idx = 1 To size
        myArray(idx) = idx
    Next idx
End Sub
";
            string expectedCode =
$@"
Option Explicit

Private myArray() As Integer

Private Sub InitializeArray(size As Long)
    Redim renamedArray(size)
    Dim idx As Long
    For idx = 1 To size
        renamedArray(idx) = idx
    Next idx
End Sub
";

            var vbe = MockVbeBuilder.BuildFromStdModules((sourceModuleName, inputCode));
            var actualModuleCode = RefactoredCode(vbe.Object, state => TestModel(state, ("myArray", "renamedArray", "renamedArray")));

            var results = actualModuleCode[sourceModuleName];
            Assert.AreEqual(expectedCode, results);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(ReplaceReferencesRefactoringAction))]
        public void ClassModuleReferences()
        {
            var sourceModuleName = "SourceModule";
            var referenceExpression = $"{sourceModuleName}.";
            var sourceModuleCode =
$@"

Public this As Long";

            string classModuleReferencingCode =
$@"Option Explicit

Private Const bar As Long = 7

Public Sub Bar()
    {referenceExpression}this = bar
End Sub

Public Sub Foo()
    With {sourceModuleName}
        .this = bar
    End With
End Sub
";

            var vbe = MockVbeBuilder.BuildFromModules((sourceModuleName, sourceModuleCode, ComponentType.StandardModule), (MockVbeBuilder.TestModuleName, classModuleReferencingCode, ComponentType.ClassModule));
            var actualModuleCode = RefactoredCode(vbe.Object, state => TestModel(state, ("this", "test", "MyProperty")));

            var referencingModuleCode = actualModuleCode[MockVbeBuilder.TestModuleName];
            StringAssert.Contains($"{sourceModuleName}.MyProperty = ", referencingModuleCode);
            StringAssert.DoesNotContain($"{sourceModuleName}.{sourceModuleName}.MyProperty = ", referencingModuleCode);
            StringAssert.Contains($"  .MyProperty = bar", referencingModuleCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(ReplaceReferencesRefactoringAction))]
        public void UDTField_MemberAccess()
        {
            var inputCode =
$@"
Private Type TBar
    FirstVal As String
    SecondVal As Long
    ThirdVal As Byte
End Type

Private myBar As TBar
Private myFoo As TBar

Public Function GetOne() As String
    GetOne = myBar.FirstVal
End Function

Public Function GetTwo() As Long
    GetTwo = myBar.ThirdVal
End Function

Public Function GetThree() As Long
    GetThree = myFoo.ThirdVal
End Function

";

            var vbe = MockVbeBuilder.BuildFromStdModules((MockVbeBuilder.TestModuleName, inputCode));
            var state = MockParser.CreateAndParse(vbe.Object);
            using (state)
            {
                var target = state.DeclarationFinder.UserDeclarations(DeclarationType.Variable)
                    .OfType<VariableDeclaration>()
                    .Where(d => "myBar" == d.IdentifierName)
                    .Single();

                var udtMembers = state.DeclarationFinder.UserDeclarations(DeclarationType.UserDefinedTypeMember)
                    .Where(d => "TBar" == d.ParentDeclaration.IdentifierName);

                var test = new UserDefinedTypeInstance(target, udtMembers);
                var refs = test.UDTMemberReferences;
                Assert.AreEqual(2, refs.Select(rf => rf.IdentifierName).Count());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(ReplaceReferencesRefactoringAction))]
        public void UDTField_MemberAccessMultipleInstances()
        {
            var inputCode =
$@"
Private Type TBar
    FirstVal As String
    SecondVal As Long
    ThirdVal As Byte
End Type

Private myBar As TBar
Private myFoo As TBar

Public Function GetOne() As String
    GetOne = myBar.FirstVal
End Function

Public Function GetTwo() As Long
    GetTwo = myBar.ThirdVal
End Function

Public Function GetThree() As Long
    GetThree = myFoo.ThirdVal
End Function

";

            var vbe = MockVbeBuilder.BuildFromStdModules((MockVbeBuilder.TestModuleName, inputCode));
            var state = MockParser.CreateAndParse(vbe.Object);
            using (state)
            {
                var myBarTarget = state.DeclarationFinder.UserDeclarations(DeclarationType.Variable)
                    .OfType<VariableDeclaration>()
                    .Where(d => "myBar" == d.IdentifierName)
                    .Single();

                var myFooTarget = state.DeclarationFinder.UserDeclarations(DeclarationType.Variable)
                    .OfType<VariableDeclaration>()
                    .Where(d => "myFoo" == d.IdentifierName)
                    .Single();

                var udtMembers = state.DeclarationFinder.UserDeclarations(DeclarationType.UserDefinedTypeMember)
                    .Where(d => "TBar" == d.ParentDeclaration.IdentifierName);

                var myBarRefs = new UserDefinedTypeInstance(myBarTarget, udtMembers);
                var refs = myBarRefs.UDTMemberReferences;
                Assert.AreEqual(2, refs.Select(rf => rf.IdentifierName).Count());

                var myFooRefs = new UserDefinedTypeInstance(myFooTarget, udtMembers);
                var fooRefs = myBarRefs.UDTMemberReferences;
                Assert.AreEqual(2, fooRefs.Select(rf => rf.IdentifierName).Count());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(ReplaceReferencesRefactoringAction))]
        public void UDTField_WithMemberAccess()
        {
            var inputCode =
$@"
Private Type TBar
    FirstVal As String
    SecondVal As Long
    ThirdVal As Byte
End Type

Private myBar As TBar
Private myFoo As TBar

Public Function GetOne() As String
    With myBar
        GetOne = .FirstVal
    End With
End Function

Public Function GetTwo() As Long
    With myBar
        GetTwo = .SecondVal
    End With
End Function

Public Function GetThree() As Long
    With myFoo
        GetThree = .ThirdVal
    End With
End Function

";

            var vbe = MockVbeBuilder.BuildFromStdModules((MockVbeBuilder.TestModuleName, inputCode));
            var state = MockParser.CreateAndParse(vbe.Object);
            using (state)
            {
                var target = state.DeclarationFinder.UserDeclarations(DeclarationType.Variable)
                    .OfType<VariableDeclaration>()
                    .Where(d => "myBar" == d.IdentifierName)
                    .Single();

                var udtMembers = state.DeclarationFinder.UserDeclarations(DeclarationType.UserDefinedTypeMember)
                    .Where(d => "TBar" == d.ParentDeclaration.IdentifierName);

                var test = new UserDefinedTypeInstance(target, udtMembers);
                var refs = test.UDTMemberReferences;
                Assert.AreEqual(2, refs.Select(rf => rf.IdentifierName).Count());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(ReplaceReferencesRefactoringAction))]
        public void ClassModuleUDTField_ExternalReferences()
        {
            var className = "TestClass";
            var classInputCode =
$@"

Public this As TBar
";

            var classInstanceName = "theClass";
            var proceduralModuleName = MockVbeBuilder.TestModuleName;
            var procedureModuleReferencingCode =
$@"Option Explicit

Public Type TBar
    First As String
    Second As Long
End Type

Private {classInstanceName} As {className}
Private Const foo As String = ""Foo""
Private Const bar As Long = 7

Public Sub Initialize()
    Set {classInstanceName} = New {className}
End Sub

Public Sub Foo()
    {classInstanceName}.this.First = foo
End Sub

Public Sub Bar()
    {classInstanceName}.this.Second = bar
End Sub

Public Sub FooBar()
    With {classInstanceName}
        .this.First = foo
        .this.Second = bar
    End With
End Sub
";

            var vbe = MockVbeBuilder.BuildFromModules((className, classInputCode, ComponentType.ClassModule),
                (proceduralModuleName, procedureModuleReferencingCode, ComponentType.StandardModule));

            var actualModuleCode = RefactoredCode(vbe.Object, state => TestModel(state, ("this", "MyType", "MyType")));

            var referencingModuleCode = actualModuleCode[proceduralModuleName];
            StringAssert.Contains($"{classInstanceName}.MyType.First = ", referencingModuleCode);
            StringAssert.Contains($"{classInstanceName}.MyType.Second = ", referencingModuleCode);
            StringAssert.Contains($"  .MyType.Second = ", referencingModuleCode);
        }

        [TestCase("DeclaringModule.this", "DeclaringModule.TheFirst.First")]
        [TestCase("this", "DeclaringModule.TheFirst.First")]
        [Category("Encapsulate Field")]
        [Category("Refactorings")]
        [Category(nameof(ReplaceReferencesRefactoringAction))]
        public void PublicUDT_ExternalFieldReferences(string memberAccessExpression, string expectedExpression)
        {
            var moduleName = "DeclaringModule";
            var inputCode =
$@"
Public this As TBazz

Public Property Let TheFirst(ByVal RHS As String)
    'this.First = RHS
End Property
";

            var referencingModuleName = MockVbeBuilder.TestModuleName;
            var referencingCode =
$@"Option Explicit

Public Type TBazz
    First As String
End Type

Public Sub Fizz()
    {memberAccessExpression}.First = ""Fizz""
End Sub
";

            var vbe = MockVbeBuilder.BuildFromModules((moduleName, inputCode, ComponentType.StandardModule),
                (referencingModuleName, referencingCode, ComponentType.StandardModule));

            var actualModuleCode = RefactoredCode(vbe.Object, state => TestModel(state, ("this", "TheFirst", "TheFirst")));

            var referencingModuleCode = actualModuleCode[referencingModuleName];
            StringAssert.Contains($"{expectedExpression} = ", referencingModuleCode);
        }

        private ReplaceReferencesModel TestModel(RubberduckParserState state, params (string fieldID, string internalName, string externalName)[] fieldConversions)
        {
            var model = new ReplaceReferencesModel()
            {
                ModuleQualifyExternalReferences = true,
            };

            foreach (var (fieldID, internalName, externalName) in fieldConversions)
            {
                var fieldDeclaration = GetUniquelyNamedDeclaration(state, DeclarationType.Variable, fieldID);
                foreach (var reference in fieldDeclaration.References)
                {
                    var replacementExpression = fieldDeclaration.QualifiedModuleName != reference.QualifiedModuleName
                        ? externalName
                        : internalName;

                    model.RegisterReferenceReplacementExpression(reference, replacementExpression);
                }
            }
            return model;
        }

        private static bool IsExternalReference(IdentifierReference identifierReference)
            => identifierReference.QualifiedModuleName != identifierReference.Declaration.QualifiedModuleName;

        private static Declaration GetUniquelyNamedDeclaration(IDeclarationFinderProvider declarationFinderProvider, DeclarationType declarationType, string identifier)
        {
            return declarationFinderProvider.DeclarationFinder.UserDeclarations(declarationType).Single(d => d.IdentifierName.Equals(identifier));
        }

        protected override IRefactoringAction<ReplaceReferencesModel> TestBaseRefactoring(RubberduckParserState state, IRewritingManager rewritingManager)
        {
            return new ReplaceReferencesRefactoringAction(rewritingManager);
        }
    }
}
