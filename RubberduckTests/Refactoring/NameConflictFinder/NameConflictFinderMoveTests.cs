using NUnit.Framework;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Common;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;
using System;
using System.Collections.Generic;
using System.Linq;
using TestDI = RubberduckTests.Refactoring.NameConflictFinderTestsDI;

namespace RubberduckTests.Refactoring
{
    [TestFixture]
    public class NameConflictFinderMoveTests
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
        [Category(nameof(NameConflictFinder))]
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
        [Category(nameof(NameConflictFinder))]
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
        [Category(nameof(NameConflictFinder))]
        public void MovedPropertyInconsistentSignatures(string letParameters, string getParameters, bool expected)
        {
            var sourceContent =
$@"
Option Explicit

Public Property Let Fi|zz{letParameters}
End Property
";
            var destinationCode =
$@"
Option Explicit

Public Property Get Fizz{getParameters} As Long
End Property
";
            Assert.AreEqual(expected, TestForMoveConflict(sourceContent, destinationCode));
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

        [TestCase("Bizz", false)]
        [TestCase("mfizZ", true)]
        [TestCase("Fizz", true)]
        [Category("Refactoring")]
        [Category(nameof(NameConflictFinder))]
        public void MoveChangedNameHasConflicts(string newName, bool expected)
        {
            var selection = ("Bazz", DeclarationType.Procedure);

            var sourceCode =
$@"
Public Sub Bazz() 
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
            Assert.AreEqual(expected, TestForMoveConflict(sourceCode, selection, destinationCode, newName: newName));
        }

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


                var proxy = MemberProxy(target, destinationModuleName, state);
                proxy.Accessibility = Accessibility.Public;
                var conflictFinder = TestDI.Resolve<INameConflictFinder>(state);
                var hasConflict = conflictFinder.ProposedDeclarationCreatesConflict(proxy, out _);
                Assert.AreEqual(isConflict, hasConflict);
            }
        }

        private bool TestForMoveConflict(string sourceContent, string destinationCode, string destinationModuleName = "DestinationDefault")
        {
            (string code, Selection selection) = ToCodeAndSelectionTuple(sourceContent);

            return TestForMoveConflict(MockVbeBuilder.TestModuleName, selection, 
                (MockVbeBuilder.TestModuleName, ComponentType.StandardModule, code),
                (destinationModuleName, ComponentType.StandardModule, destinationCode));
        }

        private bool TestForMoveConflict(string inputCode, (string identifier, DeclarationType declarationType) selection, string destinationCode, string destinationModuleName = "DestinationDefault", string newName = null)
        {
            var result = false;
            var vbe = MockVbeBuilder.BuildFromStdModules((MockVbeBuilder.TestModuleName, inputCode), (destinationModuleName, destinationCode)).Object;
            var state = MockParser.CreateAndParse(vbe);
            using (state)
            {
                var target = state.DeclarationFinder.DeclarationsWithType(selection.declarationType)
                                .Single(d => d.IdentifierName == selection.identifier && d.QualifiedModuleName.ComponentName == MockVbeBuilder.TestModuleName);

                var proxy = MemberProxy(target, destinationModuleName, state);
                proxy.IdentifierName = newName ?? target.IdentifierName;
                var conflictFinder = TestDI.Resolve<INameConflictFinder>(state);
                result = conflictFinder.ProposedDeclarationCreatesConflict(proxy, out _);
            }

            return result;
        }

        private bool TestForMoveConflict(string selectionModuleName, Selection selection, params (string moduleName, ComponentType componentType, string inputCode)[] modules)
        {
            var builder = new MockVbeBuilder()
                            .ProjectBuilder(MockVbeBuilder.TestProjectName, ProjectProtection.Unprotected);

            foreach ((string moduleName, ComponentType componentType, string inputCode) in modules)
            {
                builder = builder.AddComponent(moduleName, componentType, inputCode);
            }

            var vbe = builder.AddProjectToVbeBuilder()
                            .Build();

            var result = false;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var module = state.DeclarationFinder.AllModules.First(qmn => qmn.ComponentName.Equals(selectionModuleName));
                var qualifiedSelection = new QualifiedSelection(module, selection);
                var target = state.DeclarationFinder.AllDeclarations
                                .Where(item => item.IsUserDefined)
                                .FirstOrDefault(item => item.IsSelected(qualifiedSelection) || item.References.Any(r => r.IsSelected(qualifiedSelection)));


                var proxy = MemberProxy(target, "DestinationDefault", state);
                var conflictFinder = TestDI.Resolve<INameConflictFinder>(state);
                result = conflictFinder.ProposedDeclarationCreatesConflict(proxy, out _);
            }
            return result;
        }

        private IDeclarationProxy MemberProxy(Declaration target, string targetModuleName, RubberduckParserState state)
        {
            var targetModule = state.DeclarationFinder.MatchName(targetModuleName)
                                                .OfType<ModuleDeclaration>()
                                                .Single();

            var proxyFactory = TestDI.Resolve<IDeclarationProxyFactory>(state);
            var proxy = proxyFactory.Create(target, target.IdentifierName, targetModule);
            return proxy;
        }

        private (string code, Selection selection) ToCodeAndSelectionTuple(string input)
        {
            var codeString = input.ToCodeString();
            return (codeString.Code, codeString.CaretPosition.ToOneBased());
        }
    }

    public class NameConflictFinderTestsDI
    {
        public static T Resolve<T>(RubberduckParserState state) where T : class
        {
            return Resolve<T>(state, typeof(T).Name);
        }

        private static T Resolve<T>(RubberduckParserState _state, string name) where T : class
        {
            switch (name)
            {
                case nameof(INameConflictFinder):
                    return new NameConflictFinder(_state, Resolve<IDeclarationProxyFactory>(_state)) as T;
                case nameof(IDeclarationProxyFactory):
                    return new DeclarationProxyFactory(_state) as T;
                default:
                    throw new ArgumentException();
            }
        }
    }

}
