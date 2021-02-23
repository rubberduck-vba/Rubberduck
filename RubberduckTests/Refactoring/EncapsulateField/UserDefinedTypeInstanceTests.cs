using NUnit.Framework;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;
using System.Linq;
using Rubberduck.Refactorings.EncapsulateField;
using System.Collections.Generic;

namespace RubberduckTests.Refactoring.EncapsulateField
{
    [TestFixture]
    public class UserDefinedTypeInstanceTests
    {

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(UserDefinedTypeInstance))]
        public void GetCorrectReferenceCount()
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
";

            var vbe = MockVbeBuilder.BuildFromStdModules((MockVbeBuilder.TestModuleName, inputCode));
            var state = MockParser.CreateAndParse(vbe.Object);
            using (state)
            {
                var target = GetFieldDeclaration(state, "myBar");
                var udtMembers = GetUDTMembers(state, "TBar");

                var test = new UserDefinedTypeInstance(target, udtMembers);
                var refs = test.UDTMemberReferences;
                Assert.AreEqual(2, refs.Select(rf => rf.IdentifierName).Count());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void GetCorrectReferenceCountPerInstance()
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
                var myBarTarget = GetFieldDeclaration(state, "myBar");
                var myFooTarget = GetFieldDeclaration(state, "myFoo");
                var udtMembers = GetUDTMembers(state, "TBar");

                var myBarRefs = new UserDefinedTypeInstance(myBarTarget, udtMembers);
                var refs = myBarRefs.UDTMemberReferences;
                Assert.AreEqual(2, refs.Select(rf => rf.IdentifierName).Count());

                var myFooRefs = new UserDefinedTypeInstance(myFooTarget, udtMembers);
                var fooRefs = myFooRefs.UDTMemberReferences;
                Assert.AreEqual(1, fooRefs.Select(rf => rf.IdentifierName).Count());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void GetCorrectReferenceCount_WithMemberAccess()
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
                var target = GetFieldDeclaration(state, "myBar");

                var udtMembers = GetUDTMembers(state, "TBar");

                var test = new UserDefinedTypeInstance(target, udtMembers);
                var refs = test.UDTMemberReferences;
                Assert.AreEqual(2, refs.Select(rf => rf.IdentifierName).Count());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(UserDefinedTypeInstanceTests))]
        public void GetsCorrectReferenceCount()
        {
            string inputCode =
$@"
Private Type TBar
    First As String
    Second As String
End Type

Public Type ToEnsureValidCounts
    First As String
    Second As String
End Type

Private bizz_ As TBar

Private fizz_ As TBar

Public Sub Fizz(newValue As String)
    With bizz_
        .First = newValue
    End With
End Sub

Public Sub Buzz(newValue As String)
    With bizz_
        .Second = newValue
    End With
End Sub

Public Sub Fizz1(newValue As String)
    bizz_.First = newValue
End Sub

Public Sub Buzz1(newValue As String)
    bizz_.Second = newValue
End Sub

Public Sub Tazz(newValue As String)
    fizz_.First = newValue
    fizz_.Second = newValue
End Sub
";

            var vbe = MockVbeBuilder.BuildFromModules((MockVbeBuilder.TestModuleName, inputCode, ComponentType.StandardModule));
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var bizz_Target = state.DeclarationFinder.UserDeclarations(DeclarationType.Variable)
                    .Where(d => d.IdentifierName == "bizz_")
                    .Single();

                var udtMembers = state.DeclarationFinder.UserDeclarations(DeclarationType.UserDefinedTypeMember);

                var sut = new UserDefinedTypeInstance(bizz_Target as VariableDeclaration, udtMembers);
                Assert.AreEqual(4, sut.UDTMemberReferences.Count());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(UserDefinedTypeInstanceTests))]
        public void GetsCorrectReferenceCount_ClassAccessor()
        {
            string className = "TestClass";
            string classCode =
$@"
Public this As TBar
";

            string classInstance = "theClass";
            string moduleName = MockVbeBuilder.TestModuleName;
            string moduleCode =
$@"
Public Type TBar
    First As String
    Second As String
End Type

Public Type ToEnsureValidCounts
    First As String
    Second As String
End Type

Private {classInstance} As {className}

Public Sub Initialize()
    Set {classInstance} = New {className}
End Sub

Public Sub Fizz1(newValue As String)
        {classInstance}.this.First = newValue
End Sub

Public Sub Buzz1(newValue As String)
        {classInstance}.this.Second = newValue
End Sub

";

            var vbe = MockVbeBuilder.BuildFromModules((moduleName, moduleCode, ComponentType.StandardModule),
                (className, classCode, ComponentType.ClassModule));
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var this_Target = state.DeclarationFinder.UserDeclarations(DeclarationType.Variable)
                    .Where(d => d.IdentifierName == "this")
                    .Single();

                var udtMembers = state.DeclarationFinder.UserDeclarations(DeclarationType.UserDefinedTypeMember);

                var sut = new UserDefinedTypeInstance(this_Target as VariableDeclaration, udtMembers);
                Assert.AreEqual(2, sut.UDTMemberReferences.Count());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(UserDefinedTypeInstanceTests))]
        public void SingleElementRefNestedWithStatements()
        {
            string moduleName = MockVbeBuilder.TestModuleName;
            string moduleCode =
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

            var vbe = MockVbeBuilder.BuildFromModules((moduleName, moduleCode, ComponentType.StandardModule));

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var this_Target = state.DeclarationFinder.UserDeclarations(DeclarationType.Variable)
                    .Where(d => d.IdentifierName == "mTypesField")
                    .Single();

                var udtMembers = state.DeclarationFinder.UserDeclarations(DeclarationType.UserDefinedTypeMember);

                var sut = new UserDefinedTypeInstance(this_Target as VariableDeclaration, udtMembers);
                Assert.AreEqual(4, sut.UDTMemberReferences.Count());
            }
        }

        private static VariableDeclaration GetFieldDeclaration(IDeclarationFinderProvider declarationFinderProvider, string identifier)
        {
            return declarationFinderProvider.DeclarationFinder.UserDeclarations(DeclarationType.Variable)
                .OfType<VariableDeclaration>()
                .Where(d => identifier == d.IdentifierName)
                .Single();
        }

        private static IEnumerable<Declaration> GetUDTMembers(IDeclarationFinderProvider declarationFinderProvider, string identifier)
        {
            return declarationFinderProvider.DeclarationFinder.UserDeclarations(DeclarationType.UserDefinedTypeMember)
                .Where(d => identifier == d.ParentDeclaration.IdentifierName);
        }
    }
}
