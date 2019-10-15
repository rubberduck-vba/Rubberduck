using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Moq;
using NUnit.Framework;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using RubberduckTests.Mocks;

namespace RubberduckTests.Rewriter
{
    public class MemberAttributeRecovererTests
    {
        [Test]
        [Category("Rewriter")]
        public void RecoveringAttributesStillThereDoesNotDoAnything_ViaMember()
        {
            var inputCode =
                @"Public Function Foo() As Variant
Attribute Foo.VB_UserMemId = 0
Attribute Foo.VB_Description = ""DefaultMember""
End Function

Public Function Bar() As Variant
Attribute Bar.VB_UserMemId = -4
Attribute Bar.VB_Description = ""Enumerator""
End Function";

            var expectedCode =
                @"Public Function Foo() As Variant
Attribute Foo.VB_UserMemId = 0
Attribute Foo.VB_Description = ""DefaultMember""
End Function

Public Function Bar() As Variant
Attribute Bar.VB_UserMemId = -4
Attribute Bar.VB_Description = ""Enumerator""
End Function";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component).Object;
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                var attributesUpdater = new AttributesUpdater(state);
                var mockFailureNotifier = new Mock<IMemberAttributeRecoveryFailureNotifier>(); 
                var memberAttributeRecoverer = new MemberAttributeRecoverer(state, state, attributesUpdater, mockFailureNotifier.Object);
                memberAttributeRecoverer.RewritingManager = rewritingManager;

                var fooDeclaration = state.DeclarationFinder.UserDeclarations(DeclarationType.Function)
                    .First(decl => decl.IdentifierName.Equals("Foo"));
                var membersToRecoverAttributesFor = new List<QualifiedMemberName> {fooDeclaration.QualifiedName};

                memberAttributeRecoverer.RecoverCurrentMemberAttributesAfterNextParse(membersToRecoverAttributesFor);

                state.OnParseRequested(this);
            }
            var actualCode = component.CodeModule.Content();
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Rewriter")]
        public void RecoveringAttributesStillThereDoesNotDoAnything_ViaModule()
        {
            var inputCode =
                @"Public Function Foo() As Variant
Attribute Foo.VB_UserMemId = 0
Attribute Foo.VB_Description = ""DefaultMember""
End Function

Public Function Bar() As Variant
Attribute Bar.VB_UserMemId = -4
Attribute Bar.VB_Description = ""Enumerator""
End Function";

            var expectedCode =
                @"Public Function Foo() As Variant
Attribute Foo.VB_UserMemId = 0
Attribute Foo.VB_Description = ""DefaultMember""
End Function

Public Function Bar() As Variant
Attribute Bar.VB_UserMemId = -4
Attribute Bar.VB_Description = ""Enumerator""
End Function";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component).Object;
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                var attributesUpdater = new AttributesUpdater(state);
                var mockFailureNotifier = new Mock<IMemberAttributeRecoveryFailureNotifier>();
                var memberAttributeRecoverer = new MemberAttributeRecoverer(state, state, attributesUpdater, mockFailureNotifier.Object);
                memberAttributeRecoverer.RewritingManager = rewritingManager;

                var modulesToRecoverMemberAttributesIn = new List<QualifiedModuleName> { component.QualifiedModuleName };

                memberAttributeRecoverer.RecoverCurrentMemberAttributesAfterNextParse(modulesToRecoverMemberAttributesIn);

                state.OnParseRequested(this);
            }
            var actualCode = component.CodeModule.Content();
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Rewriter")]
        public void RecoveringAttributesRecoversTheAttributesForTheMembersProvided_ViaMember()
        {
            var inputCode =
                @"Public Function Foo() As Variant
Attribute Foo.VB_UserMemId = 0
Attribute Foo.VB_Description = ""DefaultMember""
End Function

Public Function Bar() As Variant
Attribute Bar.VB_UserMemId = -4
Attribute Bar.VB_Description = ""Enumerator""
End Function";

            var expectedCode =
                @"Public Function Foo() As Variant
Attribute Foo.VB_UserMemId = 0
Attribute Foo.VB_Description = ""DefaultMember""
End Function

Public Function Bar() As Variant
End Function";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component).Object;
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                var attributesUpdater = new AttributesUpdater(state);
                var mockFailureNotifier = new Mock<IMemberAttributeRecoveryFailureNotifier>();
                var memberAttributeRecoverer = new MemberAttributeRecoverer(state, state, attributesUpdater, mockFailureNotifier.Object);
                memberAttributeRecoverer.RewritingManager = rewritingManager;

                var fooDeclaration = state.DeclarationFinder.UserDeclarations(DeclarationType.Function)
                    .First(decl => decl.IdentifierName.Equals("Foo"));
                var membersToRecoverAttributesFor = new List<QualifiedMemberName> { fooDeclaration.QualifiedName };

                memberAttributeRecoverer.RecoverCurrentMemberAttributesAfterNextParse(membersToRecoverAttributesFor);

                var rewriteSession = rewritingManager.CheckOutCodePaneSession();
                var declarationsForWhichToRemoveAttributes = state.DeclarationFinder.UserDeclarations(DeclarationType.Function);
                RemoveAttributes(declarationsForWhichToRemoveAttributes, rewriteSession);

                ExecuteAndWaitForParserState(state, () => rewriteSession.TryRewrite(), ParserState.Ready);
            }
            var actualCode = component.CodeModule.Content();
            Assert.AreEqual(expectedCode, actualCode);
        }

        private static void ExecuteAndWaitForParserState(RubberduckParserState state, Action action, ParserState parserStateToAwait)
        {
            using (var waitHandle = new EventWaitHandle(false, EventResetMode.ManualReset))
            {
                var parserStateAwaiter = new ParserStateAwaiter(waitHandle, new List<ParserState> { parserStateToAwait });
                state.StateChanged += parserStateAwaiter.ParserStateHandler;

                action.Invoke();

                waitHandle.WaitOne();
                state.StateChanged -= parserStateAwaiter.ParserStateHandler;
            }
        }

        private static void RemoveAttributes(IEnumerable<Declaration> declarationsForWhichToRemoveAttributes, IRewriteSession rewriteSession)
        {
            foreach (var declaration in declarationsForWhichToRemoveAttributes)
            {
                foreach (var attribute in declaration.Attributes)
                {
                    //We cannot remove use the attributesUpdater here because it requires an attribute rewrite session,
                    //which rewrites in a suspended state. 
                    var attributeContext = attribute.Context;
                    var rewriter = rewriteSession.CheckOutModuleRewriter(declaration.QualifiedModuleName);
                    rewriter.Remove(attribute.Context);
                    if (attributeContext.TryGetFollowingContext(out VBAParser.EndOfLineContext followingEndOfLine))
                    {
                        rewriter.Remove(followingEndOfLine);
                    }
                }
            }
        }

        [Test]
        [Category("Rewriter")]
        public void RecoveringAttributesRecoversTheAttributesInTheModulesProvided_ViaModule()
        {
            var inputCode =
                @"Public Function Foo() As Variant
Attribute Foo.VB_UserMemId = 0
Attribute Foo.VB_Description = ""DefaultMember""
End Function

Public Function Bar() As Variant
Attribute Bar.VB_UserMemId = -4
Attribute Bar.VB_Description = ""Enumerator""
End Function";

            var expectedCodeWithRecovery =
                @"Public Function Foo() As Variant
Attribute Foo.VB_UserMemId = 0
Attribute Foo.VB_Description = ""DefaultMember""
End Function

Public Function Bar() As Variant
Attribute Bar.VB_UserMemId = -4
Attribute Bar.VB_Description = ""Enumerator""
End Function";

            var expectedCodeWithoutRecovery =
                @"Public Function Foo() As Variant
End Function

Public Function Bar() As Variant
End Function";

            var vbe = MockVbeBuilder.BuildFromStdModules(("RecoveryModule", inputCode), ("NoRecoveryModule", inputCode)).Object;
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                var attributesUpdater = new AttributesUpdater(state);
                var mockFailureNotifier = new Mock<IMemberAttributeRecoveryFailureNotifier>();
                var memberAttributeRecoverer = new MemberAttributeRecoverer(state, state, attributesUpdater, mockFailureNotifier.Object);
                memberAttributeRecoverer.RewritingManager = rewritingManager;

                var recoveryModule = state.DeclarationFinder.UserDeclarations(DeclarationType.Module)
                    .First(decl => decl.IdentifierName.Equals("RecoveryModule")).QualifiedModuleName;
                var noRecoveryModule = state.DeclarationFinder.UserDeclarations(DeclarationType.Module)
                    .First(decl => decl.IdentifierName.Equals("NoRecoveryModule")).QualifiedModuleName;

                var modulesToRecoverAttributesIn = new List<QualifiedModuleName> { recoveryModule };

                memberAttributeRecoverer.RecoverCurrentMemberAttributesAfterNextParse(modulesToRecoverAttributesIn);

                var rewriteSession = rewritingManager.CheckOutCodePaneSession();
                var declarationsForWhichToRemoveAttributes = state.DeclarationFinder.UserDeclarations(DeclarationType.Function);
                RemoveAttributes(declarationsForWhichToRemoveAttributes, rewriteSession);

                ExecuteAndWaitForParserState(state, () => rewriteSession.TryRewrite(), ParserState.Ready);

                var actualCodeWithRecovery = state.ProjectsProvider.Component(recoveryModule).CodeModule.Content();
                var actualCodeWithoutRecovery = state.ProjectsProvider.Component(noRecoveryModule).CodeModule.Content();
                Assert.AreEqual(expectedCodeWithRecovery, actualCodeWithRecovery);
                Assert.AreEqual(expectedCodeWithoutRecovery, actualCodeWithoutRecovery);
            }
        }

        [Test]
        [Category("Rewriter")]
        [TestCase("Attribute Foo.VB_UserMemId = 0", "", "")]
        [TestCase("", "Attribute Foo.VB_UserMemId = 0", "")]
        [TestCase("", "", "Attribute Foo.VB_UserMemId = 0")]
        [TestCase("Attribute Foo.VB_Description = \"myPropertyGet\"", "Attribute Foo.VB_Description = \"myPropertyLet\"", "")]
        [TestCase("", "Attribute Foo.VB_Description = \"myPropertyLet\"", "Attribute Foo.VB_Description = \"myPropertySet\"")]
        [TestCase("Attribute Foo.VB_Description = \"myPropertyGet\"", "", "Attribute Foo.VB_Description = \"myPropertySet\"")]
        [TestCase("Attribute Foo.VB_Description = \"myPropertyGet\"", "Attribute Foo.VB_Description = \"myPropertyLet\"", "Attribute Foo.VB_Description = \"myPropertySet\"")]
        public void RecoveringAttributesRecoversTheAttributesInTheModulesProvided_ViaModule_PropertiesAreHandlesSeparately(string getAttributes, string letAttributes, string setAttributes)
        {
            var inputCode =
                $@"Public Property Get Foo() As Variant
{getAttributes}
End Property

Public Property Let Foo(ByVal RHS As Long)
{letAttributes}
End Property

Public Property Set Foo(ByVal RHs As Object)
{setAttributes}
End Property";

            var expectedCodeWithRecovery = inputCode;

            var expectedCodeWithoutRecovery =
                $@"Public Property Get Foo() As Variant{(string.IsNullOrEmpty(getAttributes) ? Environment.NewLine : string.Empty)}
End Property

Public Property Let Foo(ByVal RHS As Long){(string.IsNullOrEmpty(letAttributes) ? Environment.NewLine : string.Empty)}
End Property

Public Property Set Foo(ByVal RHs As Object){(string.IsNullOrEmpty(setAttributes) ? Environment.NewLine : string.Empty)}
End Property";

            var vbe = MockVbeBuilder.BuildFromStdModules(("RecoveryModule", inputCode), ("NoRecoveryModule", inputCode)).Object;
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                var attributesUpdater = new AttributesUpdater(state);
                var mockFailureNotifier = new Mock<IMemberAttributeRecoveryFailureNotifier>();
                var memberAttributeRecoverer = new MemberAttributeRecoverer(state, state, attributesUpdater, mockFailureNotifier.Object);
                memberAttributeRecoverer.RewritingManager = rewritingManager;

                var recoveryModule = state.DeclarationFinder.UserDeclarations(DeclarationType.Module)
                    .First(decl => decl.IdentifierName.Equals("RecoveryModule")).QualifiedModuleName;
                var noRecoveryModule = state.DeclarationFinder.UserDeclarations(DeclarationType.Module)
                    .First(decl => decl.IdentifierName.Equals("NoRecoveryModule")).QualifiedModuleName;

                var modulesToRecoverAttributesIn = new List<QualifiedModuleName> { recoveryModule };

                memberAttributeRecoverer.RecoverCurrentMemberAttributesAfterNextParse(modulesToRecoverAttributesIn);

                var rewriteSession = rewritingManager.CheckOutCodePaneSession();
                var declarationsForWhichToRemoveAttributes = state.DeclarationFinder.UserDeclarations(DeclarationType.Property);
                RemoveAttributes(declarationsForWhichToRemoveAttributes, rewriteSession);

                ExecuteAndWaitForParserState(state, () => rewriteSession.TryRewrite(), ParserState.Ready);

                var actualCodeWithRecovery = state.ProjectsProvider.Component(recoveryModule).CodeModule.Content();
                var actualCodeWithoutRecovery = state.ProjectsProvider.Component(noRecoveryModule).CodeModule.Content();
                Assert.AreEqual(expectedCodeWithRecovery, actualCodeWithRecovery);
                Assert.AreEqual(expectedCodeWithoutRecovery, actualCodeWithoutRecovery);
            }
        }
    }
}