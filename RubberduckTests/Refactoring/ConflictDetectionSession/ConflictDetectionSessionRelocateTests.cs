using NUnit.Framework;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Common;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;
using System.Collections.Generic;
using System.Linq;
using TestResolver = RubberduckTests.Refactoring.ConflictDetectorTestsResolver;

namespace RubberduckTests.Refactoring
{
    [TestFixture]
    public class ConflictDetectionSessionRelocateTests
    {
        //MS_VBAL 5.3.1.6: each subroutine and Function name must be different than
        //any other module variable, module Constant, EnumerationMember, or Procedure
        //defined in the same module
        [TestCase("mFazz", DeclarationType.Function, true)]
        [TestCase("constFazz", DeclarationType.Function, true)]
        [TestCase("Bazz", DeclarationType.Function, true)]
        [TestCase("Fazz", DeclarationType.Function, true)]
        [TestCase("Fizz", DeclarationType.Function, true)]
        [TestCase("SecondValue", DeclarationType.Function, true)]
        [TestCase("Gazz", DeclarationType.Function, false)]
        [TestCase("ETest", DeclarationType.Function, false)]
        [TestCase("mFazz", DeclarationType.Procedure, true)]
        [TestCase("constFazz", DeclarationType.Procedure, true)]
        [TestCase("Bazz", DeclarationType.Procedure, true)]
        [TestCase("Fazz", DeclarationType.Procedure, true)]
        [TestCase("Fizz", DeclarationType.Procedure, true)]
        [TestCase("SecondValue", DeclarationType.Procedure, true)]
        [TestCase("Gazz", DeclarationType.Procedure, false)]
        [TestCase("ETest", DeclarationType.Procedure, false)]
        [Category("Refactoring")]
        [Category(nameof(ConflictDetectionSession))]
        public void MethodMoveConflicts(string functionName, DeclarationType declarationType, bool expected)
        {
            var methodType = declarationType.HasFlag(DeclarationType.Function)
                ? $"Function"
                : "Sub";

            var signature = declarationType.HasFlag(DeclarationType.Function)
                ? $"{functionName}() As Long"
                : $"{functionName}()";

            var selection = (functionName, declarationType);
            var sourceContent =
$@"
Option Explicit

Public {methodType} {signature}
End {methodType}
";

            var destinationCode =
$@"
Option Explicit

Public Enum ETest
    FirstValue = 0
    SecondValue
End Enum

Private mFazz As String

Private Const constFazz As Long = 7

Private mFizz As Long

Public Function Fizz() As Long
    Fizz = mFizz
End Function

Public Sub Bazz() 
End Sub

Public Property Get Fazz() As Long
    Fazz = mFazz
End Property

Public Property Let Fazz(value As Long)
    mFazz =  value
End Property
";
            Assert.AreEqual(expected, TestForMoveConflict(sourceContent, selection, destinationCode));
        }

        [TestCase("mFazz", DeclarationType.Variable, true)]
        [TestCase("constFazz", DeclarationType.Variable, true)]
        [TestCase("Bazz", DeclarationType.Variable, true)]
        [TestCase("Fazz", DeclarationType.Variable, true)]
        [TestCase("Fizz", DeclarationType.Variable, true)]
        [TestCase("SecondValue", DeclarationType.Variable, true)]
        [TestCase("Gazz", DeclarationType.Variable, false)]
        [TestCase("ETest", DeclarationType.Variable, false)]
        [TestCase("mFazz", DeclarationType.Constant, true)]
        [TestCase("constFazz", DeclarationType.Constant, true)]
        [TestCase("Bazz", DeclarationType.Constant, true)]
        [TestCase("Fazz", DeclarationType.Constant, true)]
        [TestCase("Fizz", DeclarationType.Constant, true)]
        [TestCase("SecondValue", DeclarationType.Constant, true)]
        [TestCase("Gazz", DeclarationType.Constant, false)]
        [TestCase("ETest", DeclarationType.Constant, false)]
        [Category("Refactoring")]
        [Category(nameof(ConflictDetectionSession))]
        public void VariableAndConstantMoveConflicts(string identifier, DeclarationType decType, bool expected)
        {
            var declaration = decType.HasFlag(DeclarationType.Variable)
                ? $"{identifier} As Long"
                : $"Const {identifier} As Long = 6";

            var selection = (identifier, decType);
            var sourceContent =
$@"
Option Explicit

Public {declaration}
";

            var destinationCode =
$@"
Option Explicit

Public Enum ETest
    FirstValue = 0
    SecondValue
End Enum

Private mFazz As String

Private Const constFazz As Long = 7

Private mFizz As Long

Public Function Fizz() As Long
    Fizz = mFizz
End Function

Public Sub Bazz() 
End Sub

Public Property Get Fazz() As Long
    Fazz = mFazz
End Property

Public Property Let Fazz(value As Long)
    mFazz =  value
End Property
";
            Assert.AreEqual(expected, TestForMoveConflict(sourceContent, selection, destinationCode));
        }

        //MS_VBAL 5.3.1.7: 
        //Each property Let\Set\Get must have a unique name
        [TestCase(DeclarationType.PropertyGet, "Fizz", false)]
        [TestCase(DeclarationType.PropertyLet, "Fizz", true)]
        [TestCase(DeclarationType.PropertyGet, "Fuzz", true)]
        [TestCase(DeclarationType.PropertySet, "Fuzz", false)]
        [Category("Refactoring")]
        [Category(nameof(ConflictDetectionSession))]
        public void MovedLetSetGetAreUnique(DeclarationType targetDeclarationType, string targetName, bool expected)
        {
            var selection = (targetName, targetDeclarationType);
            var sourceContent =
$@"
Option Explicit

Private mFizz As Long
Private mFuzz As Variant

Public Property Let Fizz(value As Long)
    mFizz =  value
End Property

Public Property Get Fizz() As Long
    Fizz = mFizz
End Property

Public Property Set Fuzz(var As Variant)
    Set mFuzz = Variant
End Property

Public Property Get Fuzz() As Variant
    If IsObject(mFuzz) Then
        Set Fuzz = mFuzz
    Else
        Fuzz = mFuzz
    Endif
End Property
";
            var destinationCode =
$@"
Option Explicit


Private mFizz As Long
Private mFuzz As Variant

Public Property Let Fizz(value As Long)
    mFizz =  value
End Property

Public Property Get Fuzz() As Variant
    If IsObject(mFuzz) Then
        Set Fuzz = mFuzz
    Else
        Fuzz = mFuzz
    Endif
End Property
";
            Assert.AreEqual(expected, TestForMoveConflict(sourceContent, selection, destinationCode));
        }

        //MS_VBAL 5.3.1.7: 
        //Each shared name must have the same asType, parameters, etc
        [TestCase("(value As Long)", "()", false)]
        [TestCase("(value As Variant)", "()", true)] //Inconsistent AsTypeName
        [TestCase("(value As Long)", "(arg1 As String)", true)] //Inconsistent parameters (quantity)
        [TestCase("(arg1 As Boolean, value As Long)", "(arg1 As String)", true)] //Inconsistent parameters (type)
        [TestCase("(ByVal arg1 As String, value As Long)", "(arg1 As String)", true)] //Inconsistent parameters (parameter mechanism)
        [TestCase("(arg1 As String, arg2 As Long, value As Long)", "(arg1 As String, arg22 As Long)", true)] //Inconsistent parameters (parameter name)
        [TestCase("(arg1 As String, arg2 As Variant, value As Long)", "(arg1 As String, ParamArray arg2() As Variant)", true)] //Inconsistent parameters (ParamArray)
        [Category("Refactoring")]
        [Category(nameof(ConflictDetectionSession))]
        public void MovedPropertyInconsistentSignatures(string letParameters, string getParameters, bool expected)
        {
            var selection = ("Fizz", DeclarationType.PropertyLet);
            var source =
$@"
Option Explicit

Public Property Let Fizz{letParameters}
End Property
";
            var destinationCode =
$@"
Option Explicit

Public Property Get Fizz{getParameters} As Long
End Property
";
            Assert.AreEqual(expected, TestForMoveConflict(source, selection, destinationCode));
        }

        //MS_VBAL 5.3.1.6:
        [TestCase("mFazz", true)]
        [TestCase("constFazz", true)]
        [TestCase("Fizz", true)]
        [TestCase("Bazz", true)]
        [TestCase("Fazz", true)]
        [TestCase("ETest", false)]
        [Category("Refactoring")]
        [Category(nameof(ConflictDetectionSession))]
        public void EnumMemberMoveConflicts(string enumMemberName, bool expected)
        {
            var selection = ("ETest", DeclarationType.Enumeration);

            var sourceCode =
$@"
Public Enum ETest
    {enumMemberName} = 0
    SecondValue
End Enum
";

            var destinationCode =
$@"
Option Explicit

Private mFazz As String

Private Const constFazz As Long = 7

Private mFizz As Long

Public Function Fizz() As Long
    Fizz = mFizz
End Function

Public Sub Bazz() 
End Sub

Public Property Get Fazz() As Long
    Fazz = mFazz
End Property

Public Property Let Fazz(value As Long)
    mFazz =  value
End Property
";
            Assert.AreEqual(expected, TestForMoveConflict(sourceCode, selection, destinationCode));
        }

        [TestCase("Public", "Private")]
        [TestCase("Private", "Public")]
        [Category("Refactoring")]
        [Category(nameof(ConflictDetectionSession))]
        public void EnumerationMoveConflicts(string sourceAccessibility, string destinationAccessiblity)
        {
            var selection = ("ETest", DeclarationType.Enumeration);

            var sourceCode =
$@"
{sourceAccessibility} Enum ETest
    FirstValue = 0
    SecondValue
End Enum
";

            var destinationCode =
$@"
Option Explicit

{destinationAccessiblity} Enum ETest
    SomeValue = 0
End Enum
";
            Assert.AreEqual(true, TestForMoveConflict(sourceCode, selection, destinationCode));
        }

        [TestCase("Public", "Private")]
        [TestCase("Private", "Public")]
        [Category("Refactoring")]
        [Category(nameof(ConflictDetectionSession))]
        public void UserDefinedTypeMoveConflicts(string sourceAccessibility, string destinationAccessiblity)
        {
            var selection = ("TTest", DeclarationType.UserDefinedType);

            var sourceCode =
$@"
{sourceAccessibility} Type TTest
    FirstValue As Long
    SecondValue As String
End Type
";

            var destinationCode =
$@"
Option Explicit

{destinationAccessiblity} Type TTest
    SomeValue As Boolean
End Type
";
            Assert.AreEqual(true, TestForMoveConflict(sourceCode, selection, destinationCode));
        }

        [TestCase(true, false)]
        [TestCase(false, true)]
        [Category(nameof(ConflictDetectionSession))]
        public void MovedPrivateToPublicConstant(bool useModuleQualification, bool isConflict)
        {
            var relocateSourceModuleName = MockVbeBuilder.TestModuleName;
            var relocateSourceModuleContent =
$@"
Option Explicit

Private Const THE_CONST As Long = 4

Public Property Get TheConst() As Long
    TheConst = THE_CONST
End Property
";
            var destinationModuleCode =
$@"
Option Explicit
";

            var conflictModuleName = "ConflictModule";
            var moduleQualification = useModuleQualification ? $"{conflictModuleName}." : string.Empty;
            var conflictModuleCode =
$@"
Option Explicit

Public Const THE_CONST As Long = 6

Public Function TimesSix(arg As Long) As Long
    TimesSix = arg * THE_CONST
End Function
";
            var conflictReferenceModuleName = "ConflictReferenceModule";
            var conflictReferenceModuleCode =
$@"
Option Explicit

Public Function TimesSixty(arg As Long) As Long
    TimesSix = arg * {moduleQualification}THE_CONST * 10
End Function
";
            var destinationModuleName = "DestinationModule";

            var vbe = MockVbeBuilder.BuildFromModules(
                (relocateSourceModuleName, relocateSourceModuleContent, ComponentType.StandardModule),
                (destinationModuleName, destinationModuleCode, ComponentType.StandardModule),
                (conflictModuleName, conflictModuleCode, ComponentType.StandardModule),
                (conflictReferenceModuleName, conflictReferenceModuleCode, ComponentType.StandardModule)
                );

            var state = MockParser.CreateAndParse(vbe.Object);
            using (state)
            {
                var target = state.DeclarationFinder.DeclarationsWithType(DeclarationType.Constant)
                    .Where(d => d.QualifiedModuleName.ComponentName.Equals(MockVbeBuilder.TestModuleName)).Single();

                var destinationModule = state.DeclarationFinder.DeclarationsWithType(DeclarationType.Module)
                    .Where(d => d.IdentifierName.Equals(destinationModuleName)).OfType<ModuleDeclaration>().Single();

                var conflictSession = TestResolver.Resolve<IConflictSessionFactory>(state).Create();
                var relocateConflictDetector = conflictSession.RelocateConflictDetector;
                var hasConflict = relocateConflictDetector.IsConflictingName(target, destinationModule, out _, Accessibility.Public);

                Assert.AreEqual(isConflict, hasConflict);
            }
        }

        [TestCase("Bar", DeclarationType.PropertyLet)]
        [TestCase("Bar", DeclarationType.PropertyGet)]
        [TestCase("Bar", DeclarationType.PropertySet)]
        [Category("Refactorings")]
        [Category(nameof(ConflictDetectionSession))]
        public void DoesNotConflictMovedPropertyLSGWithUnMovedLSG(string targetIdentifier, DeclarationType targetDeclarationType)
        {
            var selection = (targetIdentifier, targetDeclarationType);
            var source = $@"
Option Explicit

Private mBar As Variant

Public Property Let Bar(arg1 As Variant)
    mBar = arg1
End Property

Public Property Set Bar(arg1 As Variant)
    Set mBar = arg1
End Property

Public Property Get Bar() As Variant
    If IsObject(mBar) Then
        Set Bar = mBar
    Else
        Bar = mBar
    End If
End Property
";
            var destinationCode =
$@"
Option Explicit

Public Function Foo(arg1 As Variant) As Variant
    If IsObject(arg1) Then
        Set arg1 = Bar
        Set Bar = arg1
        Set Foo = arg1
    Else
        arg1 = Bar
        Bar = arg1
        Foo = arg1
    End If
End Function

";
            Assert.AreEqual("Bar", TestForMovedNonConflictNamePairs(source, selection, destinationCode));
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(ConflictDetectionSession))]
        public void PropertyGetWithPropertyLetOfDifferentParameters()
        {
            var selection = ("Bar", DeclarationType.PropertyLet);
            var source = $@"
Option Explicit

Private mBar As Long

Public Property Let Bar(arg1 As Long)
    mBar = arg1
End Property

Public Property Get Bar() As Long
    Bar = mBar
End Property
";
            var destinationCode =
$@"
Option Explicit

Private mMyBar As Long

Public Function Foo(arg1 As Long) As Long
    arg1 = Bar * 10
    {MockVbeBuilder.TestModuleName}.Bar = arg1
    Foo = arg1
End Function

Public Property Get Bar(arg1 As Long) As Long
    Bar = mMyBar
End Property
";
            Assert.AreEqual("Bar1", TestForMovedNonConflictNamePairs(source, selection, destinationCode));
        }

        [Test]
        [Category("Refactoring")]
        [Category(nameof(ConflictDetectionSession))]
        public void RespectsPreviousRenamesAndExistingIdentifiers()
        {
            var sourceCode =
$@"
Option Explicit

Private mTestVar1 As Long

Private mTestVar2 As Long

Private mTestVar3 As Long

Private SameName2 As Long

";

            var FieldsToRename = new string[] { "mTestVar1", "mTestVar2", "mTestVar3" };

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(sourceCode, out _);
            var state = MockParser.CreateAndParse(vbe.Object);
            using (state)
            {
                var targets = state.DeclarationFinder.DeclarationsWithType(DeclarationType.Variable)
                                .Where(d => d.IdentifierName.StartsWith("mTestVar") && d.QualifiedModuleName.ComponentName == MockVbeBuilder.TestModuleName);

                var nonConflictNames = new List<string>();
                var factory = TestResolver.Resolve<IConflictSessionFactory>(state);
                var conflictSession = factory.Create();
                foreach (var target in targets)
                {
                    var targetProxy = conflictSession.ProxyCreator.Create(target, "SameName");
                    conflictSession.TryRegister(targetProxy, out _, true);
                }
                var renamePairs = conflictSession.RenamePairs;

                StringAssert.AreEqualIgnoringCase("SameName", renamePairs.ElementAt(0).NewName);
                StringAssert.AreEqualIgnoringCase("SameName1", renamePairs.ElementAt(1).NewName);
                StringAssert.AreEqualIgnoringCase("SameName3", renamePairs.ElementAt(2).NewName);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(ConflictDetectionSession))]
        public void MovePrivateEnumRespectsDestinationNameCollision()
        {
            var selection = ("KeyOne", DeclarationType.EnumerationMember);
            var sourceCode =
$@"
Option Explicit

Private Enum KeyValues
    KeyOne
    KeyTwo
End Enum
";

            var destinationCode =
$@"
Option Explicit

Private Sub KeyOne(arg As Long)
End Sub
";
            Assert.AreEqual("KeyOne1", TestForMovedNonConflictNamePairs(sourceCode, selection, destinationCode));
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(ConflictDetectionSession))]
        public void MoveEnumAllowsSameMemberNames()
        {
            var selection = ("KeyValues", DeclarationType.Enumeration);
            var sourceCode =
$@"
Option Explicit

Private Enum KeyValues
    KeyOne
    KeyTwo
End Enum
";

            var destinationCode =
$@"
Option Explicit

Private Enum KeyValues2
    KeyOne
    KeyTwo
End Enum
";
            var results = TestForMovedNonConflictNames(sourceCode, selection, destinationCode);
            Assert.AreEqual(0, results.Count());
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(ConflictDetectionSession))]
        public void MovePrivateEnumResolvesNameCollisions()
        {
            var selection = ("KeyValues", DeclarationType.Enumeration);
            var sourceCode =
$@"
Option Explicit

Private Enum KeyValues
    KeyOne
    KeyTwo
End Enum
";

            var destinationCode =
$@"
Option Explicit

Private keyOne1 As Double
Private keyOne2 As Double
Private keyOne3 As Double
Private keyOne5 As Double

Private Const keyTWO1 As String = ""Fizz""
Private Const keyTWO2 As String = ""Fuzz""

Private Sub KeyOne(arg As Long)
End Sub

Private Sub KeyTwo(arg As Long)
End Sub
";
            var renamePairs = TestForMovedNonConflictNames(sourceCode, selection, destinationCode);
            Assert.AreEqual(2, renamePairs.Count());
            Assert.IsTrue(renamePairs.Select(r => r.newName).Contains("KeyOne4"));
            Assert.IsTrue(renamePairs.Select(r => r.newName).Contains("KeyTwo3"));
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(ConflictDetectionSession))]
        public void Multiple()
        {
            var selection = ("KeyValues", DeclarationType.Enumeration);
            var sourceCode =
$@"
Option Explicit

Private Enum KeyValues
    KeyOne
    KeyTwo
End Enum
";

            var destinationCode =
$@"
Option Explicit

Private keyOne1 As Double
Private keyOne2 As Double
Private keyOne3 As Double
Private keyOne5 As Double

Private Const keyTWO1 As String = ""Fizz""
Private Const keyTWO2 As String = ""Fuzz""

Private Sub KeyOne(arg As Long)
End Sub

Private Sub KeyTwo(arg As Long)
End Sub
";

            var renamingPairs = TestForMovedNonConflictNames(sourceCode, selection, destinationCode);
            Assert.AreEqual(2, renamingPairs.Count());
        }

        [TestCase("Bizz", false)]
        [TestCase("mfizZ", true)]
        [TestCase("Fizz", true)]
        [Category("Refactoring")]
        [Category(nameof(ConflictDetectionSession))]
        public void MoveChangedNameHasConflicts(string memberName, bool expected)
        {
            var selection = (memberName, DeclarationType.Procedure);

            var sourceCode =
$@"
Public Sub {memberName}() 
End Sub
";

            var destinationCode =
$@"
Option Explicit

Private mFizz As Long

Public Function Fizz() As Long
    Fizz = mFizz
End Function
";
            Assert.AreEqual(expected, TestForMoveConflict(sourceCode, selection, destinationCode));
        }

        [TestCase("fizz", "fizz1")]
        [TestCase("fizz1", "fizz2")]
        [TestCase("fizz123", "fizz124")]
        [TestCase("f67oo3", "f67oo4")]
        [TestCase("fizz0", "fizz1")]
        [TestCase("", "1")]
        [Category("Refactorings")]
        [Category(nameof(ConflictDetectionSession))]
        public void IdentifierNameIncrementing(string input, string expected)
        {
            var actual = ConflictDetectorBase.IncrementIdentifier(input);
            Assert.AreEqual(expected, actual);
        }

        private bool TestForMoveConflict(string inputCode, (string identifier, DeclarationType declarationType) selection, string destinationCode, string destinationModuleName = "DestinationDefault")
        {
            var isConflict = false;
            var nonConflictName = string.Empty;
            var vbe = MockVbeBuilder.BuildFromStdModules((MockVbeBuilder.TestModuleName, inputCode), (destinationModuleName, destinationCode)).Object;
            var state = MockParser.CreateAndParse(vbe);
            using (state)
            {
                var target = state.DeclarationFinder.DeclarationsWithType(selection.declarationType)
                                .Single(d => d.IdentifierName == selection.identifier && d.QualifiedModuleName.ComponentName == MockVbeBuilder.TestModuleName);

                var destinationModule = state.DeclarationFinder.DeclarationsWithType(DeclarationType.ProceduralModule)
                                .OfType<ModuleDeclaration>()
                                .Single(d => d.QualifiedModuleName.ComponentName == destinationModuleName);

                var conflictDetector = TestResolver.Resolve<IConflictSessionFactory>(state).Create().RelocateConflictDetector;
                isConflict = conflictDetector.IsConflictingName(target, destinationModule, out _);
            }

            return isConflict;
        }

        private string TestForMovedNonConflictNamePairs(string inputCode, (string identifier, DeclarationType declarationType) selection, string destinationCode, string destinationModuleName = "DestinationDefault")
        {
            var results = TestForMovedNonConflictNames(inputCode, selection, destinationCode, destinationModuleName);
            return results.Any() ? results.First().newName : selection.identifier;
        }

        private List<(Declaration target, string newName)> TestForMovedNonConflictNames(string inputCode, (string identifier, DeclarationType declarationType) selection, string destinationCode, string destinationModuleName = "DestinationDefault")
        {
            var results = new List<(Declaration target, string newName)>();
            var vbe = MockVbeBuilder.BuildFromStdModules((MockVbeBuilder.TestModuleName, inputCode), (destinationModuleName, destinationCode)).Object;
            var state = MockParser.CreateAndParse(vbe);
            using (state)
            {
                var target = state.DeclarationFinder.DeclarationsWithType(selection.declarationType)
                                .Single(d => d.IdentifierName == selection.identifier && d.QualifiedModuleName.ComponentName == MockVbeBuilder.TestModuleName);

                var destinationModule = state.DeclarationFinder.DeclarationsWithType(DeclarationType.ProceduralModule)
                                .Single(d => d.IdentifierName == destinationModuleName) as ModuleDeclaration;

                var conflictSession = TestResolver.Resolve<IConflictSessionFactory>(state).Create();

                var proxy = conflictSession.ProxyCreator.Create(target);
                proxy.TargetModule = destinationModule;
                conflictSession.TryRegisterRelocation(proxy, out _, true);
                results = conflictSession.RenamePairs.ToList();
            }

            return results;
        }
    }
}
