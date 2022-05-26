using NUnit.Framework;
using Rubberduck.Refactorings.EncapsulateField;
using RubberduckTests.Mocks;
using Rubberduck.Refactorings;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.Utility;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor.SafeComWrappers;
using System.Linq;

namespace RubberduckTests.Refactoring.EncapsulateField
{
    [TestFixture]
    public class EncapsulateFieldValidatorTests : EncapsulateFieldInteractiveRefactoringTest
    {
        private EncapsulateFieldTestSupport Support { get; } = new EncapsulateFieldTestSupport();

        [TestCase("fizz", "_Fizz", false)]
        [TestCase("fizz", "FizzProp", true)]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void VBAIdentifier_Property(string originalFieldName, string newPropertyName, bool expectedResult)
        {
            var inputCode =
$@"Public {originalFieldName} As String";

            var encapsulatedField = Support.RetrieveEncapsulateFieldCandidate(inputCode, originalFieldName);

            encapsulatedField.EncapsulateFlag = true;
            encapsulatedField.PropertyIdentifier = newPropertyName;
            Assert.AreEqual(expectedResult, encapsulatedField.TryValidateEncapsulationAttributes(out _));
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulatePrivateField_ReadOnlyRequiresSet()
        {
            var inputCode =
                @"|Private fizz As Collection";

            const string expectedCode =
                @"Private fizz As Collection

Public Property Get Name() As Collection
    Set Name = fizz
End Property
";
            var presenterAction = Support.SetParametersForSingleTarget("fizz", "Name", isReadonly: true);
            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            Assert.AreEqual(expectedCode.Trim(), actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void PropertyNameNotDuplicated()
        {
            var inputCode =
                @"Public var|iable As Integer, variable1 As Long, variable2 As Integer";

            var userInput = new UserInputDataObject()
                .UserSelectsField("variable")
                .UserSelectsField("variable1")
                .UserSelectsField("variable2");

            var presenterAction = Support.SetParameters(userInput);
            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            StringAssert.Contains("Public Property Get Variable() As Integer", actualCode);
            StringAssert.Contains("Variable = variable3", actualCode);
            StringAssert.Contains("Public Property Get Variable1() As Long", actualCode);
            StringAssert.Contains("Variable1 = variable4", actualCode);
            StringAssert.Contains("Public Property Get Variable2() As Integer", actualCode);
            StringAssert.Contains("Variable2 = variable5", actualCode);
            StringAssert.DoesNotContain("Public Property Get Variable3() As Integer", actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UDTMemberPropertyConflictsWithExistingFunction()
        {
            var inputCode =
$@"
Private Type TBar
    First As String
    Second As Long
End Type

Private myBar As TBar

Private Function First() As String
    First = myBar.First
End Function";

            var candidate = Support.RetrieveEncapsulateFieldCandidate(inputCode, "myBar", DeclarationType.Variable);
            var result = candidate.ConflictFinder.IsConflictingIdentifier(candidate, "First", out _);
            Assert.AreEqual(true, result);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void FieldNameDefaultsToNonConflictName()
        {
            var inputCode =
$@"Public fizz As String

            Private fizzle As String

            'fizz1 is the initial default name for encapsulating 'fizz'            
            Public Property Get Fizz1() As String
                Fizz1 = fizzle
            End Property

            Public Property Let Fizz1(ByVal value As String)
                fizzle = value
            End Property
            ";
            var encapsulatedField = Support.RetrieveEncapsulateFieldCandidate(inputCode, "fizz");
            Assert.IsTrue(encapsulatedField.TryValidateEncapsulationAttributes(out _));
        }

        [TestCase("Name")]
        [TestCase("mName")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UserEntersConflictingName(string userModifiedPropertyName)
        {
            var inputCode =
$@"Public fizz As String

            Private mName As String

            Public Property Get Name() As String
                Name = mName
            End Property

            Public Property Let Name(ByVal value As String)
                mName = value
            End Property
            ";

            var fieldUT = "fizz";
            var userInput = new UserInputDataObject()
                .UserSelectsField("fizz", userModifiedPropertyName);

            var presenterAction = Support.SetParameters(userInput);

            var model = Support.RetrieveUserModifiedModelPriorToRefactoring(inputCode, fieldUT, DeclarationType.Variable, presenterAction);

            Assert.IsFalse(model["fizz"].TryValidateEncapsulationAttributes(out _));
        }

        [TestCase("Number", "Bazzle", true, true)]
        [TestCase("Number", "Number", false, false)]
        [TestCase("Test", "Number", false, true)]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UserModificationIsExistingPropertyNameConflicts(string fizz_modifiedPropertyName, string bazz_modifiedPropertyName, bool fizz_expectedResult, bool bazz_expectedResult)
        {
            var inputCode =
$@"Public fizz As Integer
Public bazz As Integer
Public buzz As Integer

Private mTest As Integer

Public Property Get Test() As Integer
    Test = mTest
End Property";

            var fieldUT = "fizz";
            var userInput = new UserInputDataObject()
                .UserSelectsField(fieldUT, fizz_modifiedPropertyName, true)
                .UserSelectsField("bazz", bazz_modifiedPropertyName, true);

            var presenterAction = Support.SetParameters(userInput);

            var model = Support.RetrieveUserModifiedModelPriorToRefactoring(inputCode, fieldUT, DeclarationType.Variable, presenterAction);

            Assert.AreEqual(fizz_expectedResult, model["fizz"].TryValidateEncapsulationAttributes(out _), "fizz failed");
            Assert.AreEqual(bazz_expectedResult, model["bazz"].TryValidateEncapsulationAttributes(out _), "bazz failed");
        }

        [TestCase("Private", "Private")]
        [TestCase("Public", "Private")]
        [TestCase("Private", "Public")]
        [TestCase("Public", "Public")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulateMultipleUDTFields_DefaultsAreNotInConflict(string udtAccessibility, string fieldAccessibility)
        {
            var inputCode =
$@"
{udtAccessibility} Type TBar
    First As Long
    Second As String
End Type

{fieldAccessibility} this As TBar

{fieldAccessibility} that As TBar
";
            var fieldUT = "this";
            var userInput = new UserInputDataObject()
                .UserSelectsField(fieldUT)
                .UserSelectsField("that");

            var presenterAction = Support.SetParameters(userInput);

            var model = Support.RetrieveUserModifiedModelPriorToRefactoring(inputCode, fieldUT, DeclarationType.Variable, presenterAction);

            Assert.AreEqual(true, model[fieldUT].TryValidateEncapsulationAttributes(out var message), message);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void PropertyNameConflictsWithModuleVariable()
        {
            var inputCode =
$@"
Public longValue As Long

Public wholeNumber As String
";
            var fieldUT = "longValue";
            var userInput = new UserInputDataObject()
                .UserSelectsField(fieldUT, "WholeNumber", true);

            var presenterAction = Support.SetParameters(userInput);

            var model = Support.RetrieveUserModifiedModelPriorToRefactoring(inputCode, fieldUT, DeclarationType.Variable, presenterAction);

            Assert.AreEqual(false, model[fieldUT].TryValidateEncapsulationAttributes(out _));
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulatePrivateField_EnumMemberConflict()
        {
            //5.2.3.4: An enum member name may not be the same as any variable name, or constant name that is defined within the same module
            var inputCode =
                @"

Public Enum NumberTypes 
     Whole = -1 
     Integral = 0 
     Rational1 = 1 
End Enum

Private rati|onal As NumberTypes
";

            var presenterAction = Support.UserAcceptsDefaults();
            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            StringAssert.Contains("Public Property Get Rational() As NumberTypes", actualCode);
            StringAssert.Contains("Rational = rational2", actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulatePrivateField_UDTMemberConflict()
        {
            var inputCode =
                @"

Private Type TVehicle
    Wheels As Integer
    MPG As Double
End Type

Private whe|els As Integer
";

            var presenterAction = Support.UserAcceptsDefaults();
            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            StringAssert.Contains("Public Property Get Wheels()", actualCode);
            StringAssert.Contains("Wheels = wheels1", actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void DefaultPropertyNameConflictsResolved()
        {
            //Both fields default to "Test" as the property name
            var inputCode =
                @"Private mTest As Integer
Private strTest As String";

            var fieldUT = "mTest";

            var presenterAction = Support.UserAcceptsDefaults(fieldUT, "strTest");

            var model = Support.RetrieveUserModifiedModelPriorToRefactoring(inputCode, fieldUT, DeclarationType.Variable, presenterAction);

            Assert.AreEqual(true, model[fieldUT].TryValidateEncapsulationAttributes(out var errorMsg), errorMsg);
        }

        [TestCase("Dim test As String", "arg")] //Local variable
        [TestCase(@"Const test As String = ""Foo""", "arg")] //Local constant
        [TestCase(@"Const localTest As String = ""Foo""", "test")] //parameter
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void TargetNameUsedForLimitedScopeDeclarations(string localDeclaration, string parameter)
        {
            var inputCode =
$@"
Private te|st As Long

Private Function Foo({parameter} As String) As String
    {localDeclaration}
    test = test & ""Foo""
    Foo = test
End Function
";
            var presenterAction = Support.UserAcceptsDefaults();
            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            StringAssert.Contains("Test", actualCode);
            StringAssert.Contains("test1", actualCode);
            StringAssert.DoesNotContain("Test1", actualCode);
        }

        [TestCase("Dim test As String", "arg")] //Local variable
        [TestCase(@"Const test As String = ""Foo""", "arg")] //Local constant
        [TestCase(@"Const localTest As String = ""Foo""", "test")] //parameter
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void TargetReferenceScopeUsesPropertyName(string localDeclaration, string parameter)
        {
            var inputCode =
$@"
Private aName As String

Private Function Foo({parameter} As String) As String
    {localDeclaration}
    test = aName & test
    Foo = test
End Function
";
            var fieldUT = "aName";
            var userInput = new UserInputDataObject()
                .UserSelectsField(fieldUT, "Test", true);

            var presenterAction = Support.SetParameters(userInput);

            var model = Support.RetrieveUserModifiedModelPriorToRefactoring(inputCode, fieldUT, DeclarationType.Variable, presenterAction);

            Assert.AreEqual(false, model[fieldUT].TryValidateEncapsulationAttributes(out _));
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void TargetDefaultFieldIDConflict()
        {
            var inputCode =
$@"
Private tes|t As String
Private test1 As String

Public Sub Foo(arg As String)
    test = arg & test1
End Sub
";
            var presenterAction = Support.UserAcceptsDefaults();
            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            StringAssert.Contains("Test", actualCode);
            StringAssert.Contains("Private test2 As String", actualCode);
            StringAssert.DoesNotContain("test1 = arg & test1", actualCode);
        }

        [TestCase(MockVbeBuilder.TestModuleName)]
        [TestCase("TestModule")]
        [TestCase("TestClass")]
        [TestCase(MockVbeBuilder.TestProjectName)]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void ModuleAndProjectNamesAreValid(string userEnteredName)
        {
            var fieldUT = "foo";
            var userInput = new UserInputDataObject()
                .UserSelectsField(fieldUT, userEnteredName, true);

            var presenterAction = Support.SetParameters(userInput);

            var vbe = MockVbeBuilder.BuildFromModules(
                (MockVbeBuilder.TestModuleName, "Private foo As String", ComponentType.StandardModule),
                ("TestModule", "Private foo1 As String", ComponentType.StandardModule),
                ("TestClass", "Private foo2 As String", ComponentType.ClassModule)).Object;

            var model = Support.RetrieveUserModifiedModelPriorToRefactoring(vbe, fieldUT, DeclarationType.Variable, presenterAction);

            Assert.AreEqual(true, model[fieldUT].TryValidateEncapsulationAttributes(out _));
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void ExistingPublicUDTConflictWithDefaultObjectStateType()
        {
            string moduleOneName = "ModuleOne";
            var inputCode =
$@"
Option Explicit

Public mF|oo As Long
";

            string module2Content =
$@"
Public Type TModuleOne
    FirstVal As Long
    SecondVal As String
End Type
";

            var fieldUT = "mFoo";
            var userInput = new UserInputDataObject()
                .UserSelectsField(fieldUT);

            userInput.EncapsulateUsingUDTField();

            var presenterAction = Support.SetParameters(userInput);

            var actualModuleCode = Support.RefactoredCode(presenterAction, 
                (moduleOneName, inputCode.ToCodeString(), ComponentType.StandardModule),
                ("Module2", module2Content, ComponentType.StandardModule));

            StringAssert.Contains($"Private Type TModuleOne", actualModuleCode[moduleOneName]);
        }

        [TestCase("Public")]
        [TestCase("Private")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void ExistingUDTConflictWithDefaultObjectStateType(string accessibility)
        {
            var moduleOneName = "ModuleOne";
            var inputCode =
$@"
Option Explicit

{accessibility} Type TModuleOne
    FirstVal As Long
    SecondVal As String
End Type

Public mF|oo As Long
";

            var fieldUT = "mFoo";
            var userInput = new UserInputDataObject()
                .UserSelectsField(fieldUT);

            userInput.EncapsulateUsingUDTField();

            var presenterAction = Support.SetParameters(userInput);

            var actualModuleCode = Support.RefactoredCode(presenterAction, 
                (moduleOneName, inputCode.ToCodeString(), ComponentType.StandardModule));

            StringAssert.Contains($"Private Type TModuleOne1", actualModuleCode[moduleOneName]);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UDTReservedMemberArrayIdentifier()
        {
            var fieldName = "Name";
            var userInput = new UserInputDataObject()
                .UserSelectsField(fieldName);

            userInput.ConvertFieldsToUDTMembers = true;

            var presenterAction = Support.SetParameters(userInput);

            var vbe = MockVbeBuilder.BuildFromModules(
                (MockVbeBuilder.TestModuleName, $"Private {fieldName}(5) As String", ComponentType.StandardModule)).Object;

            var model = Support.RetrieveUserModifiedModelPriorToRefactoring(vbe, fieldName, DeclarationType.Variable, presenterAction);

            Assert.AreEqual(false, model[fieldName].TryValidateEncapsulationAttributes(out var errorMessage), errorMessage);

            var expectedMessage = string.Format(RefactoringsUI.InvalidNameCriteria_IsReservedKeywordFormat, fieldName);
            Assert.AreEqual(expectedMessage, errorMessage);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UserEntersUDTMemberPropertyNameInConflictWithExistingField()
        {
            var inputCode =
                @"

Private Type TVehicle
    Wheels As Integer
    MPG As Double
End Type

Private vehicle As TVehicle

Private seats As Integer

Private foo As String
";
            var userInput = new UserInputDataObject()
                .UserSelectsField("seats", "Foo");

            userInput.ConvertFieldsToUDTMembers = true;
            userInput.ObjectStateUDTTargetID = "TVehicle";

            var presenterAction = Support.SetParameters(userInput);

            var model = Support.RetrieveUserModifiedModelPriorToRefactoring(inputCode, "seats", DeclarationType.Variable, presenterAction);
            Assert.AreEqual(false, model["seats"].TryValidateEncapsulationAttributes(out var errorMessage), errorMessage);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UserClearsConflictingNameByEncapsulatingConflictingVariable()
        {
            var inputCode =
                @"

Private Type TVehicle
    Wheels As Integer
    MPG As Double
End Type

Private mVehicle As TVehicle

Private seats As Integer

Private foo As String
";
            //By encapsulating variable mVehicle, the variable disappears and
            //is converted to UDTMember name "Vehicle" and "Vehicle" properties
            // - thus removing the conflict created by the user editing the "seats" property
            var userInput = new UserInputDataObject()
                .UserSelectsField("seats", "MVehicle")
                .UserSelectsField("mVehicle");

            userInput.ConvertFieldsToUDTMembers = true;

            var presenterAction = Support.SetParameters(userInput);

            var model = Support.RetrieveUserModifiedModelPriorToRefactoring(inputCode, "seats", DeclarationType.Variable, presenterAction);
            Assert.AreEqual(true, model["seats"].TryValidateEncapsulationAttributes(out var errorMessage), errorMessage);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void AddedUDTMemberConflictsWithExistingName()
        {
            var fieldUT = "mFirstValue";
            var inputCode =
                $@"

Private Type MyType
    FirstValue As Integer
    SecondValue As Integer
End Type

Private {fieldUT} As Double

Private myType As MyType
";
            var userInput = new UserInputDataObject()
                .UserSelectsField(fieldUT);

            userInput.EncapsulateUsingUDTField("myType");

            var presenterAction = Support.SetParameters(userInput);

            var model = Support.RetrieveUserModifiedModelPriorToRefactoring(inputCode, fieldUT, DeclarationType.Variable, presenterAction);
            Assert.AreEqual(false, model[fieldUT].TryValidateEncapsulationAttributes(out var errorMessage), errorMessage);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void AddedFieldConflictsWithExistingUDTMemberName()
        {
            var fieldUT = "mFirstValue";
            var inputCode =
                $@"

Private Type MyType
    FirstValue As Integer
    SecondValue As Integer
End Type

Private {fieldUT} As Double

Private myType As MyType
";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _).Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                var mTypeTarget = state.DeclarationFinder.DeclarationsWithType(DeclarationType.Variable)
                    .First(d => d.IdentifierName == "myType");

                var mFirstTarget = state.DeclarationFinder.DeclarationsWithType(DeclarationType.Variable)
                    .First(d => d.IdentifierName == fieldUT);

                var resolver = Support.SetupResolver(state);

                var candidateSetsProviderFactory = resolver.Resolve<IEncapsulateFieldCandidateSetsProviderFactory>();
                var candidateSets = candidateSetsProviderFactory.Create(state, resolver.Resolve<IEncapsulateFieldCandidateFactory>(), mFirstTarget.QualifiedModuleName);

                var encapsulateFieldCandidates = candidateSets.EncapsulateFieldUseBackingFieldCandidates;

                var finderFactory = resolver.Resolve<IEncapsulateFieldConflictFinderFactory>();

                var conflictFinder = finderFactory.Create(state, candidateSets.EncapsulateFieldUseBackingFieldCandidates, candidateSets.ObjectStateFieldCandidates);

                foreach (var candidate in encapsulateFieldCandidates)
                {
                    candidate.ConflictFinder = conflictFinder;
                }

                var mTypeCandidate = encapsulateFieldCandidates.Single(c => c.Declaration == mTypeTarget);
                mTypeCandidate.EncapsulateFlag = true;

                var mFirstCandidate = encapsulateFieldCandidates.Single(c => c.Declaration == mFirstTarget);

                foreach (var candidate in encapsulateFieldCandidates)
                {
                    candidate.EncapsulateFlag = true;
                }

                var result = mFirstCandidate.TryValidateEncapsulationAttributes(out var errorMessage);
                Assert.IsTrue(result, errorMessage);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void ObjectStateUDTFieldConflictsWithAssignedProperty()
        {
            var fieldUT = "mFirstValue";
            var inputCode =
                $@"

Private {fieldUT} As Double
";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _).Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                var mFirstTarget = state.DeclarationFinder.DeclarationsWithType(DeclarationType.Variable)
                    .First(d => d.IdentifierName == fieldUT) as VariableDeclaration;

                var resolver = Support.SetupResolver(state);
                var modelFactory = resolver.Resolve<IEncapsulateFieldModelFactory>();
                var model = modelFactory.Create(mFirstTarget);
                var mFirstCandidate = model[mFirstTarget.IdentifierName];

                mFirstCandidate.EncapsulateFlag = true;
                mFirstCandidate.PropertyIdentifier = "This";

                model.EncapsulateFieldStrategy = EncapsulateFieldStrategy.ConvertFieldsToUDTMembers;
                var objectStateUDT = model.ObjectStateUDTField;

                model.ConflictFinder.AssignNoConflictIdentifiers(objectStateUDT);

                StringAssert.AreEqualIgnoringCase("this1", objectStateUDT.IdentifierName);
            }
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
