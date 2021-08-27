using Antlr4.Runtime;
using NUnit.Framework;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.DeleteDeclarations;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RubberduckTests.Refactoring.DeleteDeclarations
{
    [TestFixture]
    public class DeleteDeclarationsPropertyMembersTests
    {
        [TestCase(DeclarationType.PropertyGet)]
        [TestCase(DeclarationType.PropertyLet)]
        [TestCase(DeclarationType.PropertySet)]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void IsGroupedWithRelatedProperties_ReturnsTrue(DeclarationType declarationType)
        {
            var identifier = "TestProperty";
            var inputCode =
$@"
Public Sub Fizz()
End Sub

Public Property Get TestProperty() As Variant
End Property
Public Property Let TestProperty(ByVal RHS As Variant)
End Property
Public Property Set TestProperty(ByVal RHS As Variant)
End Property

Public Property Get SomeOtherProperty() As String
End Property
";

            var expected = true;

            void thisTest(IPropertyDeletionTarget deleteTarget)
            {
                var actual = deleteTarget.IsGroupedWithRelatedProperties();

                Assert.AreEqual(actual, expected);
            }

            SetupAndInvokeTest(inputCode, (identifier, declarationType), thisTest);
        }

        [TestCase(DeclarationType.PropertyGet)]
        [TestCase(DeclarationType.PropertyLet)]
        [TestCase(DeclarationType.PropertySet)]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void IsGroupedWithRelatedProperties_ReturnsFalse(DeclarationType declarationType)
        {
            var identifier = "TestProperty";
            var inputCode =
$@"
Public Property Get TestProperty() As Variant
End Property

Public Sub Fizz()
End Sub


Public Property Let TestProperty(ByVal RHS As Variant)
End Property

Public Property Get SomeOtherProperty() As String
End Property
Public Property Set TestProperty(ByVal RHS As Variant)
End Property
";

            var expected = false;

            void thisTest(IPropertyDeletionTarget deleteTarget)
            {
                var actual = deleteTarget.IsGroupedWithRelatedProperties();

                Assert.AreEqual(actual, expected);
            }

            SetupAndInvokeTest(inputCode, (identifier, declarationType), thisTest);
        }

        [TestCase(DeclarationType.PropertyGet, false)]
        [TestCase(DeclarationType.PropertyLet, true)]
        [TestCase(DeclarationType.PropertySet, true)]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void IsGroupedWithRelatedProperties_PairAndSingle(DeclarationType declarationType, bool expected)
        {
            var identifier = "TestProperty";
            var inputCode =
$@"
Public Property Get TestProperty() As Variant
End Property

Public Sub Fizz()
End Sub


Public Property Let TestProperty(ByVal RHS As Variant)
End Property
Public Property Set TestProperty(ByVal RHS As Variant)
End Property

Public Property Get SomeOtherProperty() As String
End Property
";

            void thisTest(IPropertyDeletionTarget deleteTarget)
            {
                var actual = deleteTarget.IsGroupedWithRelatedProperties();

                Assert.AreEqual(actual, expected);
            }

            SetupAndInvokeTest(inputCode, (identifier, declarationType), thisTest);
        }

        [TestCase(DeclarationType.PropertyGet, false)]
        [TestCase(DeclarationType.PropertyLet, false)]
        [TestCase(DeclarationType.PropertySet, false)]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void IsGroupedWithRelatedProperties_SingleProperty(DeclarationType declarationType, bool expected)
        {
            var identifier = "TestProperty";
            var declarationLine = string.Empty;
            switch (declarationType)
            {
                case DeclarationType.PropertyGet:
                    declarationLine = $"Property Get {identifier}() As Variant";
                    break;
                case DeclarationType.PropertyLet:
                    declarationLine = $"Property Let {identifier}(ByVal RHS As Variant)";
                    break;
                case DeclarationType.PropertySet:
                    declarationLine = $"Property Set {identifier}(ByVal RHS As Variant)";
                    break;
            }
            var inputCode =
$@"
Public Sub Fizz()
End Sub


Public {declarationLine}
End Property

Public Property Get SomeOtherProperty() As String
End Property
";

            void thisTest(IPropertyDeletionTarget deleteTarget)
            {
                var actual = deleteTarget.IsGroupedWithRelatedProperties();

                Assert.AreEqual(actual, expected);
            }

            SetupAndInvokeTest(inputCode, (identifier, declarationType), thisTest);
        }

        [TestCase("", "\r\n")]
        [TestCase("\r\n", "\r\n\r\n")]
        [TestCase("\r\n\r\n", "\r\n\r\n\r\n")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void RemoveLastProperty(string separation, string expectedEOSContent)
        {
            var inputCode =
$@"
Public Property Get TestProperty() As Variant
End Property
Public Property Let TestProperty(ByVal RHS As Variant)
End Property
Public Property Set TestProperty(ByVal RHS As Variant)
End Property
{separation}Public Property Get SomeOtherProperty() As String
End Property
";

            void thisTest(DeleteDeclarationsTestsResolver resolver, IDeclarationFinderProvider declarationFinderProvider, IRewritingManager rewritingManager)
            {
                var target = declarationFinderProvider.DeclarationFinder.UserDeclarations(DeclarationType.PropertySet).First(d => d.IdentifierName == "TestProperty");

                var precedingNonDeleteTarget = declarationFinderProvider.DeclarationFinder.UserDeclarations(DeclarationType.PropertyGet).First(d => d.IdentifierName == "TestProperty");

                var rewriteSession = rewritingManager.CheckOutCodePaneSession();

                var deletionTargetFactory = resolver.Resolve<IDeclarationDeletionTargetFactory>();
                var sut = deletionTargetFactory.Create(target, rewriteSession) as IModuleElementDeletionTarget;

                precedingNonDeleteTarget.Context.TryGetFollowingContext<VBAParser.EndOfStatementContext>(out var eos);

                Assert.IsNotNull(eos);
                sut.PrecedingEOSContext = eos;

                var actualEOS = sut.BuildEOSReplacementContent();

                StringAssert.AreEqualIgnoringCase(expectedEOSContent, actualEOS);
            }

            SetupAndInvokeTest(inputCode, thisTest);
        }

        [TestCase("", "\r\n")]
        [TestCase("\r\n", "\r\n\r\n")]
        [TestCase("\r\n\r\n", "\r\n\r\n\r\n")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void RemoveFirstProperty(string separation, string expectedEOSContent)
        {
            var inputCode =
$@"
Public Sub Fizz()
End Sub
{separation}Public Property Get TestProperty() As Variant
End Property
Public Property Let TestProperty(ByVal RHS As Variant)
End Property
Public Property Set TestProperty(ByVal RHS As Variant)
End Property

Public Property Get SomeOtherProperty() As String
End Property
";

            void thisTest(DeleteDeclarationsTestsResolver resolver, IDeclarationFinderProvider declarationFinderProvider, IRewritingManager rewritingManager)
            {
                var target = declarationFinderProvider.DeclarationFinder.UserDeclarations(DeclarationType.PropertyGet).First(d => d.IdentifierName == "TestProperty");

                var rewriteSession = rewritingManager.CheckOutCodePaneSession();

                var deletionTargetFactory = resolver.Resolve<IDeclarationDeletionTargetFactory>();
                var sut = deletionTargetFactory.Create(target, rewriteSession) as IModuleElementDeletionTarget;

                var precedingNonDeleteTarget = declarationFinderProvider.DeclarationFinder.MatchName("Fizz").First();

                (precedingNonDeleteTarget.Context as ParserRuleContext).TryGetFollowingContext<VBAParser.EndOfStatementContext>(out var eos);

                Assert.IsNotNull(eos);
                sut.PrecedingEOSContext = eos;

                var actualEOS = sut.BuildEOSReplacementContent();

                StringAssert.AreEqualIgnoringCase(expectedEOSContent, actualEOS);
            }

            SetupAndInvokeTest(inputCode, thisTest);
        }

        private static void SetupAndInvokeTest(string inputCode, (string ID, DeclarationType DeclarationType) targetSpec, Action<IPropertyDeletionTarget> testSUT)
        {
            var vbe = MockVbeBuilder.BuildFromModules((MockVbeBuilder.TestModuleName, inputCode, ComponentType.StandardModule)).Object;
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                var resolver = new DeleteDeclarationsTestsResolver(state, rewritingManager);

                var target = state.DeclarationFinder.UserDeclarations(targetSpec.DeclarationType).First(d => d.IdentifierName == targetSpec.ID);

                var rewriteSession = rewritingManager.CheckOutCodePaneSession();

                var deletionTargetFactory = resolver.Resolve<IDeclarationDeletionTargetFactory>();
                var deleteTarget = deletionTargetFactory.Create(target, rewriteSession) as IPropertyDeletionTarget;
                testSUT(deleteTarget);
            }
        }
        private static void SetupAndInvokeTest(string inputCode, Action<DeleteDeclarationsTestsResolver, IDeclarationFinderProvider, IRewritingManager> testSUT)
        {
            var vbe = MockVbeBuilder.BuildFromModules((MockVbeBuilder.TestModuleName, inputCode, ComponentType.StandardModule)).Object;
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                var resolver = new DeleteDeclarationsTestsResolver(state, rewritingManager);
                testSUT(resolver, state, rewritingManager);
            }
        }
    }
}
