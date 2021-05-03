using NUnit.Framework;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.ReplacePrivateUDTMemberReferences;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;
using System.Linq;

namespace RubberduckTests.Refactoring.ReplacePrivateUDTMemberReferences
{
    [TestFixture]
    public class UserDefinedTypeInstanceTests
    {
        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(ReplacePrivateUDTMemberReferencesRefactoringAction))]
        [Category(nameof(UserDefinedTypeInstanceTests))]
        public void GetsCorrectReferenceCount()
        {
            string inputCode =
$@"
Private Type TBar
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
        [Category(nameof(ReplacePrivateUDTMemberReferencesRefactoringAction))]
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

Private {classInstance} As {className}

Public Sub Initialize()
    Set {classInstance} = New {className}
End Sub

'Public Sub Fizz(newValue As String)
'    With {classInstance}
'        .this.First = newValue
'    End With
'End Sub

'Public Sub Buzz(newValue As String)
'    With {classInstance}
'        .this.Second = newValue
'    End With
'End Sub

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
    }
}
