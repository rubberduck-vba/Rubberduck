using System.Linq;
using NUnit.Framework;
using Rubberduck.Parsing.Symbols;
using RubberduckTests.Mocks;

namespace RubberduckTests.Annotations
{
    [TestFixture]
    public class AnnotationResolutionTests
    {
        [Test]
        public void MemberAnnotationsAboveMemberGetScopedToMember()
        {
            const string inputCode =
                @"
'@TestMethod
'@Enumerator 17, 12 @DefaultMember
Public Sub Foo()
End Sub

Public Function Bar() As Variant
End Function";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var fooDeclaration = state.DeclarationFinder.UserDeclarations(DeclarationType.Procedure).Single();

                var expectedAnnotationCount = 3;
                var actualAnnotationCount = fooDeclaration.Annotations.Count();

                Assert.AreEqual(expectedAnnotationCount, actualAnnotationCount);
            }
        }

        [Test]
        public void NonMemberAnnotationsAboveMemberDoNotGetScopedToMember()
        {
            const string inputCode =
                @"
'@TestModule
Public Sub Foo()
End Sub

Public Function Bar() As Variant
End Function";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var fooDeclaration = state.DeclarationFinder.UserDeclarations(DeclarationType.Procedure).Single();

                var expectedAnnotationCount = 0;
                var actualAnnotationCount = fooDeclaration.Annotations.Count();

                Assert.AreEqual(expectedAnnotationCount, actualAnnotationCount);
            }
        }

        [Test]
        public void LineContinuedMemberAnnotationsAboveMemberGetScopedToMember()
        {
            const string inputCode =
                @"
Public Sub Foo()
End Sub
'@TestMethod _

'@Enumerator _
17 _
, _
12 _
 _
@DefaultMember _

Public Function Bar() As Variant
End Function";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var barDeclaration = state.DeclarationFinder.UserDeclarations(DeclarationType.Function).Single();

                var expectedAnnotationCount = 3;
                var actualAnnotationCount = barDeclaration.Annotations.Count();

                Assert.AreEqual(expectedAnnotationCount, actualAnnotationCount);
            }
        }

        [Test]
        public void MemberAnnotationsAboveFirstNonAnnotationLineAboveMemberDoNotGetScopedToMember()
        {
            const string inputCode =
                @"
Public Sub Foo()
End Sub
'@TestMethod
'SomeComment
'@Enumerator _
17 _
, _
12 _
 _
@DefaultMember _

Public Function Bar() As Variant
End Function";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var barDeclaration = state.DeclarationFinder.UserDeclarations(DeclarationType.Function).Single();

                var expectedAnnotationCount = 2;
                var actualAnnotationCount = barDeclaration.Annotations.Count();

                Assert.AreEqual(expectedAnnotationCount, actualAnnotationCount);
            }
        }

        [Test]
        public void MemberAnnotationsAboveFirstNonMemberNonIdentifierAnnotationLineAboveMemberDoNotGetScopedToMember()
        {
            const string inputCode =
                @"
Public Sub Foo()
End Sub
'@TestMethod
'@TestModule
'@Enumerator _
17 _
, _
12 _
 _
@DefaultMember _

Public Function Bar() As Variant
End Function";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var barDeclaration = state.DeclarationFinder.UserDeclarations(DeclarationType.Function).Single();

                var expectedAnnotationCount = 2;
                var actualAnnotationCount = barDeclaration.Annotations.Count();

                Assert.AreEqual(expectedAnnotationCount, actualAnnotationCount);
            }
        }

        [Test]
        [Ignore("We cannot test this because we do not have any identifier annotation that is not a member annotation.")]
        public void MemberAnnotationsAboveIdentifierAnnotationLineAboveMemberGetScopedToMember()
        {
            const string inputCode =
                @"
Public Sub Foo()
End Sub
'@TestMethod
'@TestIdentifier
'@Enumerator _
17 _
, _
12 _
 _
@DefaultMember _

Public Function Bar() As Variant
End Function";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var barDeclaration = state.DeclarationFinder.UserDeclarations(DeclarationType.Function).Single();

                var expectedAnnotationCount = 3;
                var actualAnnotationCount = barDeclaration.Annotations.Count();

                Assert.AreEqual(expectedAnnotationCount, actualAnnotationCount);
            }
        }

        [Test]
        public void ModuleAnnotationsAboveNonAnnotationLineAboveFirstMemberAreModuleAnnotations()
        {
            const string inputCode =
                @"
Public Foobar As Long
'@TestModule
'@Folder(""Test"")
'SomeComment
'@Enumerator 17, _
12 _
@DefaultMember
Public Sub Foo()
End Sub

Public Function Bar() As Variant
End Function";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var moduleDeclaration = state.DeclarationFinder.UserDeclarations(DeclarationType.ProceduralModule).Single();

                var expectedAnnotationCount = 2;
                var actualAnnotationCount = moduleDeclaration.Annotations.Count();

                Assert.AreEqual(expectedAnnotationCount, actualAnnotationCount);
            }
        }

        [Test]
        public void NonModuleAnnotationsAboveNonAnnotationLineAboveFirstMemberAreNotModuleAnnotations()
        {
            const string inputCode =
                @"
Public Foobar As Long
'@TestMethod
'SomeComment
'@Enumerator 17, _
12 _
@DefaultMember
Public Sub Foo()
End Sub

Public Function Bar() As Variant
End Function";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var moduleDeclaration = state.DeclarationFinder.UserDeclarations(DeclarationType.ProceduralModule).Single();

                var expectedAnnotationCount = 0;
                var actualAnnotationCount = moduleDeclaration.Annotations.Count();

                Assert.AreEqual(expectedAnnotationCount, actualAnnotationCount);
            }
        }

        [Test]
        public void ModuleAnnotationsBelowFirstMemberAreNotModuleAnnotations()
        {
            const string inputCode =
                @"
Public Foobar As Long

Public Sub Foo()
End Sub

'@TestModule

Public Function Bar() As Variant
End Function";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var moduleDeclaration = state.DeclarationFinder.UserDeclarations(DeclarationType.ProceduralModule).Single();

                var expectedAnnotationCount = 0;
                var actualAnnotationCount = moduleDeclaration.Annotations.Count();

                Assert.AreEqual(expectedAnnotationCount, actualAnnotationCount);
            }
        }

        [Test]
        public void AllModuleAnnotationsAreModuleAnnotationsIfThereIsNoBody()
        {
            const string inputCode =
                @"
'@TestModule
'@Folder(""Test"")
'SomeComment
'@Enumerator 17, _
12 _
@DefaultMember
";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var moduleDeclaration = state.DeclarationFinder.UserDeclarations(DeclarationType.ProceduralModule).Single();

                var expectedAnnotationCount = 2;
                var actualAnnotationCount = moduleDeclaration.Annotations.Count();

                Assert.AreEqual(expectedAnnotationCount, actualAnnotationCount);
            }
        }

        [Test]
        public void VariableAnnotationsAboveVariableGetScopedToVariable()
        {
            const string inputCode =
                @"
'@Obsolete
Public foo As Long

Public Function Bar() As Variant
End Function";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var fooDeclaration = state.DeclarationFinder.UserDeclarations(DeclarationType.Variable).Single();

                var expectedAnnotationCount = 1;
                var actualAnnotationCount = fooDeclaration.Annotations.Count();

                Assert.AreEqual(expectedAnnotationCount, actualAnnotationCount);
            }
        }

        [Test]
        public void NonVariableAnnotationsAboveVariableDoNotGetScopedToVariable()
        {
            const string inputCode =
                @"
'@TestModule
Public foo As Long

Public Function Bar() As Variant
End Function";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var fooDeclaration = state.DeclarationFinder.UserDeclarations(DeclarationType.Variable).Single();

                var expectedAnnotationCount = 0;
                var actualAnnotationCount = fooDeclaration.Annotations.Count();

                Assert.AreEqual(expectedAnnotationCount, actualAnnotationCount);
            }
        }

        [Test]
        public void LineContinuedVariableAnnotationsAboveVariableGetScopedToVariable()
        {
            const string inputCode =
                @"
'@Obsolete _

Public foo As Long

Public Function Bar() As Variant
End Function";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var fooDeclaration = state.DeclarationFinder.UserDeclarations(DeclarationType.Variable).Single();

                var expectedAnnotationCount = 1;
                var actualAnnotationCount = fooDeclaration.Annotations.Count();

                Assert.AreEqual(expectedAnnotationCount, actualAnnotationCount);
            }
        }

        [Test]
        public void VariableAnnotationsAboveFirstNonAnnotationLineAboveVariableDoNotGetScopedToVariable()
        {
            const string inputCode =
                @"
'@Obsolete
'SomeComment
Public foo As Long

Public Function Bar() As Variant
End Function";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var barDeclaration = state.DeclarationFinder.UserDeclarations(DeclarationType.Variable).Single();

                var expectedAnnotationCount = 0;
                var actualAnnotationCount = barDeclaration.Annotations.Count();

                Assert.AreEqual(expectedAnnotationCount, actualAnnotationCount);
            }
        }

        [Test]
        public void VariableAnnotationsAboveFirstNonVariableNonIdentifierAnnotationLineAboveVariableDoNotGetScopedToVariable()
        {
            const string inputCode =
                @"
'@Obsolete
'@TestModule
Public foo As Long

Public Function Bar() As Variant
End Function";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var fooDeclaration = state.DeclarationFinder.UserDeclarations(DeclarationType.Variable).Single();

                var expectedAnnotationCount = 0;
                var actualAnnotationCount = fooDeclaration.Annotations.Count();

                Assert.AreEqual(expectedAnnotationCount, actualAnnotationCount);
            }
        }

        [Test]
        [Ignore("We cannot test this because we do not have any identifier annotation that is not a member annotation.")]
        public void VariableAnnotationsAboveIdentifierAnnotationLineAboveVariableGetScopedToVariable()
        {
            const string inputCode =
                @"
'@Obsolete
'@TestIdentifier
Public foo As Long

Public Function Bar() As Variant
End Function";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var fooDeclaration = state.DeclarationFinder.UserDeclarations(DeclarationType.Variable).Single();

                var expectedAnnotationCount = 1;
                var actualAnnotationCount = fooDeclaration.Annotations.Count();

                Assert.AreEqual(expectedAnnotationCount, actualAnnotationCount);
            }
        }
    }
}