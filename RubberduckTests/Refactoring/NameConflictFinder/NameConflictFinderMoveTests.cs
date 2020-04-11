using NUnit.Framework;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.Common;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;
using System;
using System.Collections.Generic;
using System.Linq;

namespace RubberduckTests.Refactoring
{
    [TestFixture]
    public class NameConflictFinderMoveTests
    {
        [TestCase("ConflictModule.", false)]
        [TestCase("", true)]
        [Category(nameof(NameConflictFinder))]
        public void MovedPrivateToPublicConstant(string moduleQualification, bool isConflict)
        {
            var containingModuleContent =
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

            var conflictModuleCode =
$@"
Option Explicit

Public Const THE_CONST As Long = 6

Public Function TimesSix(arg As Long) As Long
    TimesSix = arg * THE_CONST
End Function
";
            var conflictReferenceModuleCode =
$@"
Option Explicit

Public Function TimesSixty(arg As Long) As Long
    TimesSix = arg * {moduleQualification}THE_CONST * 10
End Function
";
            var renameTargetModuleName = MockVbeBuilder.TestModuleName;
            var destinationModuleName = "DestinationModule";
            var conflictModuleName = "ConflictModule";
            var conflictReferenceModuleName = "ConflictRefernceModule";

            var vbe = MockVbeBuilder.BuildFromModules(
                (renameTargetModuleName, containingModuleContent, ComponentType.StandardModule),
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
                    .Where(d => d.IdentifierName.Equals(destinationModuleName)).Single();


                var conflictFinder = new NameConflictFinder(state);
                var hasConflict = conflictFinder.MoveCreatesNameConflict(target, destinationModule.IdentifierName, Accessibility.Public, out _);
                Assert.AreEqual(isConflict, hasConflict);
            }
        }

        //MS_VBAL 5.3.1.6: each subroutine and Function name must be different than
        //any other module variable, module Constant, EnumerationMember, or Procedure
        //defined in the same module
        [TestCase("mFazz", true)]
        [TestCase("constFazz", true)]
        [TestCase("Bazz", true)]
        [TestCase("Fazz", true)]
        [TestCase("Fizz", true)]
        [TestCase("SecondValue", true)]
        [TestCase("Gazz", false)]
        [TestCase("ETest", false)]
        [Category("Refactoring")]
        [Category(nameof(NameConflictFinder))]
        public void FunctionMoveConflicts(string functionName, bool expected)
        {
            var selection = (functionName, DeclarationType.Function);
            var sourceContent =
$@"
Option Explicit

Public Function {functionName}(arg As Long) As Long
End Function
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

        //MS_VBAL 5.3.1.6:
        [TestCase("mFazz", true)]
        [TestCase("constFazz", true)]
        [TestCase("Bazz", true)]
        [TestCase("Fazz", true)]
        [TestCase("Fizz", true)]
        [TestCase("SecondValue", true)]
        [TestCase("Gazz", false)]
        [TestCase("ETest", false)]
        [Category("Refactoring")]
        [Category(nameof(NameConflictFinder))]
        public void SubroutineMoveConflicts(string subroutineName, bool expected)
        {
            var selection = (subroutineName, DeclarationType.Procedure);
            var sourceContent =
$@"
Option Explicit

Public Sub {subroutineName}(arg As Long)
End Sub
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


        //MS_VBAL 5.3.1.6:
        [TestCase("mFazz", "Public", "", DeclarationType.Variable, true)]
        [TestCase("constFazz", "Public", "", DeclarationType.Variable, true)]
        [TestCase("Bazz", "Public", "", DeclarationType.Variable, true)]
        [TestCase("Fazz", "Public", "", DeclarationType.Variable, true)]
        [TestCase("Fizz", "Public", "", DeclarationType.Variable, true)]
        [TestCase("SecondValue", "Public", "", DeclarationType.Variable, true)]
        [TestCase("Gazz", "Public", "", DeclarationType.Variable, false)]
        [TestCase("ETest", "Public", "", DeclarationType.Variable, false)]
        [TestCase("mFazz", "Public Const", "= 6", DeclarationType.Constant, true)]
        [TestCase("constFazz", "Public Const", "= 6", DeclarationType.Constant, true)]
        [TestCase("Bazz", "Public Const", "= 6", DeclarationType.Constant, true)]
        [TestCase("Fazz", "Public Const", "= 6", DeclarationType.Constant, true)]
        [TestCase("Fizz", "Public Const", "= 6", DeclarationType.Constant, true)]
        [TestCase("SecondValue", "Public Const", "= 6", DeclarationType.Constant, true)]
        [TestCase("Gazz", "Public Const", "= 6", DeclarationType.Constant, false)]
        [TestCase("ETest", "Public Const", "= 6", DeclarationType.Constant, false)]
        [Category("Refactoring")]
        [Category(nameof(NameConflictFinder))]
        public void VariableAndConstantMoveConflicts(string identifier, string declarationPrefix, string declarationSuffix, DeclarationType decType, bool expected)
        {
            var selection = (identifier, decType);
            var sourceContent =
$@"
Option Explicit

{declarationPrefix} {identifier} As Long {declarationSuffix}
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
        [Category(nameof(NameConflictFinder))]
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


Private mFizz As Variant
Private mFuzz As Variant

Public Property Let Fizz(value As Variant)
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

        //MS_VBAL 5.3.1.6:
        [TestCase("mFazz", true)]
        [TestCase("constFazz", true)]
        [TestCase("Fizz", true)]
        [TestCase("Bazz", true)]
        [TestCase("Fazz", true)]
        [TestCase("ETest", false)]
        [Category("Refactoring")]
        [Category(nameof(NameConflictFinder))]
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
        [Category(nameof(NameConflictFinder))]
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
        [Category(nameof(NameConflictFinder))]
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

        private bool TestForMoveConflict(string inputCode, (string identifier, DeclarationType declarationType) selection, string destinationCode, string destinationModuleName = "DestinationDefault")
        {
            var result = false;
            var vbe = MockVbeBuilder.BuildFromStdModules((MockVbeBuilder.TestModuleName, inputCode), (destinationModuleName, destinationCode)).Object;
            var state = MockParser.CreateAndParse(vbe);
            using (state)
            {
                var target = state.DeclarationFinder.DeclarationsWithType(selection.declarationType)
                                .Single(d => d.IdentifierName == selection.identifier && d.QualifiedModuleName.ComponentName == MockVbeBuilder.TestModuleName);

                var conflictFinder = new NameConflictFinder(state);
                result = conflictFinder.MoveCreatesNameConflict(target, destinationModuleName, target.Accessibility, out _);
            }

            return result;
        }
    }
}
