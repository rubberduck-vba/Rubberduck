using NUnit.Framework;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.Refactorings.MoveMember;
using RubberduckTests.Mocks;
using System.Linq;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using System.Collections.Generic;

namespace RubberduckTests.Refactoring.MoveMember
{
    [TestFixture]
    public class SingleFieldOrConstantTests : MoveMemberTestsBase
    {
        [TestCase(MoveEndpoints.FormToStd, "Private")]
        [TestCase(MoveEndpoints.ClassToStd, "Private")]
        [TestCase(MoveEndpoints.StdToStd, "Public")]
        [TestCase(MoveEndpoints.StdToStd, "Private")]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void XToStd_MoveConstant_HasReferences(MoveEndpoints moveEndpoints, string accessibility)
        {
            var source =
$@"
Option Explicit

{accessibility} Const mFoo As Long = 10

Public Function FooMath(arg1 As Long) As Long
    FooMath = arg1 * mFoo * 2
End Function
";

            var moveDefinition = new TestMoveDefinition(moveEndpoints, ("mFoo", DeclarationType.Constant), sourceContent: source);

            var callSiteModuleName = "CallSiteModule";
            if (moveDefinition.IsStdModuleSource && accessibility.Equals(Tokens.Public))
            {
                var otherModuleReference =
    $@"
Option Explicit

Public Function MemberAccessFoo(arg1 As Long) As Long
    MemberAccessFoo = arg1 * {moveDefinition.SourceModuleName}.mFoo * 2
End Function

Public Function WithMemberAccessFoo(arg1 As Long) As Long
    Dim result As Long
    With {moveDefinition.SourceModuleName}
        result = (.mFoo + arg1) * 2
    End With
    WithMemberAccessFoo = result
End Function

Public Function NonQualifiedFoo(arg1 As Long) As Long
    NonQualifiedFoo = (mFoo + arg1) * 2
End Function
";
                moveDefinition.Add(new ModuleDefinition(callSiteModuleName, ComponentType.StandardModule, otherModuleReference));
            }

            var destinationDeclaration = "Public Const mFoo As Long = 10";

            var refactoredCode = RefactoredCode(moveDefinition, source);

            StringAssert.DoesNotContain("Const mFoo As Long = 10", refactoredCode.Source);
            StringAssert.Contains(destinationDeclaration, refactoredCode.Destination);
            StringAssert.Contains($"{moveDefinition.DestinationModuleName}.{moveDefinition.SelectedElement}", refactoredCode.Source);
            if (moveDefinition.IsStdModuleSource && accessibility.Equals(Tokens.Public))
            {
                StringAssert.Contains($"{moveDefinition.DestinationModuleName}.{moveDefinition.SelectedElement}", refactoredCode[callSiteModuleName]);
                StringAssert.Contains($"result = ({moveDefinition.DestinationModuleName}.{moveDefinition.SelectedElement} + arg1) * 2", refactoredCode[callSiteModuleName]);
                StringAssert.Contains($"NonQualifiedFoo = ({moveDefinition.DestinationModuleName}.{moveDefinition.SelectedElement} + arg1) * 2", refactoredCode[callSiteModuleName]);
            }
        }

        [TestCase(MoveEndpoints.FormToStd, "Private")]
        [TestCase(MoveEndpoints.ClassToStd, "Private")]
        [TestCase(MoveEndpoints.StdToStd, "Public")]
        [TestCase(MoveEndpoints.StdToStd, "Private")]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void XToStd_MoveNonAggregateValueTypeField_HasReferences(MoveEndpoints moveEndpoints, string accessibility)
        {
            var source =
$@"
Option Explicit

{accessibility} mFoo As Long

Public Function FooMath(arg1 As Long) As Long
    FooMath = arg1 * mFoo * 2
End Function
";

            var moveDefinition = new TestMoveDefinition(moveEndpoints, ("mFoo", DeclarationType.Variable), sourceContent: source);

            var callSiteModuleName = "CallSiteModule";
            if (moveDefinition.IsStdModuleSource && accessibility.Equals(Tokens.Public))
            {
                var otherModuleReference =
    $@"
Option Explicit

Public Function MemberAccessFoo(arg1 As Long) As Long
    MemberAccessFoo = arg1 * {moveDefinition.SourceModuleName}.mFoo * 2
End Function

Public Function WithMemberAccessFoo(arg1 As Long) As Long
    Dim result As Long
    With {moveDefinition.SourceModuleName}
        result = (.mFoo + arg1) * 2
    End With
    WithMemberAccessFoo = result
End Function

Public Function NonQualifiedFoo(arg1 As Long) As Long
    NonQualifiedFoo = (mFoo + arg1) * 2
End Function
";
                moveDefinition.Add(new ModuleDefinition(callSiteModuleName, ComponentType.StandardModule, otherModuleReference));
            }

            var refactoredCode = RefactoredCode(moveDefinition, source);

            var destinationDeclaration = "Public mFoo As Long";

            StringAssert.DoesNotContain("mFoo As Long", refactoredCode.Source);
            StringAssert.Contains(destinationDeclaration, refactoredCode.Destination);
            StringAssert.Contains($"{moveDefinition.DestinationModuleName}.{moveDefinition.SelectedElement}", refactoredCode.Source);
            if (moveDefinition.IsStdModuleSource && accessibility.Equals(Tokens.Public))
            {
                StringAssert.Contains($"{moveDefinition.DestinationModuleName}.{moveDefinition.SelectedElement}", refactoredCode[callSiteModuleName]);
                StringAssert.Contains($"result = ({moveDefinition.DestinationModuleName}.{moveDefinition.SelectedElement} + arg1) * 2", refactoredCode[callSiteModuleName]);
                StringAssert.Contains($"NonQualifiedFoo = ({moveDefinition.DestinationModuleName}.{moveDefinition.SelectedElement} + arg1) * 2", refactoredCode[callSiteModuleName]);
            }
        }

        [TestCase("Public")]
        [TestCase("Private")]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void StdToStd_PublicUDT_MoveField_HasReferences(string accessibility)
        {
            var source =
$@"
Option Explicit

Public Type MyTestType
    Foo As Long
    Bar As String
End Type

{accessibility} mFooBar As MyTestType

Public Function FooMath(arg1 As Long) As Long
    FooMath = arg1 * mFooBar.Foo
End Function

Public Function ConcatBar(arg1 As String) As String
    ConcatBar = arg1 & mFooBar.Bar
End Function";

            var moveDefinition = new TestMoveDefinition(MoveEndpoints.StdToStd, ("mFooBar", DeclarationType.Variable), sourceContent: source);

            var callSiteModuleName = "CallSiteModule";
            if (moveDefinition.IsStdModuleSource && accessibility.Equals(Tokens.Public))
            {
                var otherModuleReference =
    $@"
Option Explicit

Public Function MemberAccessFoo(arg1 As Long) As Long
    MemberAccessFoo = arg1 + {moveDefinition.SourceModuleName}.mFooBar.Foo * 2
End Function

Public Function WithMemberAccessFoo(arg1 As Long) As Long
    Dim result As Long
    With {moveDefinition.SourceModuleName}
        result = (.mFooBar.Foo + arg1) * 2
    End With
    WithMemberAccessFoo = result
End Function

Public Function WithMemberAccessFoo2(arg1 As Long) As Long
    Dim result As Long
    With {moveDefinition.SourceModuleName}.mFooBar
        result2 = (.Foo + arg1) * 2
    End With
    WithMemberAccessFoo2 = result2
End Function

Public Function NonQualifiedFoo(arg1 As Long) As Long
    NonQualifiedFoo = (mFooBar.Foo + arg1) * 2
End Function
";
                moveDefinition.Add(new ModuleDefinition(callSiteModuleName, ComponentType.StandardModule, otherModuleReference));
            }

            var refactoredCode = RefactoredCode(moveDefinition, source);

            var destinationDeclaration = "Public mFooBar As MyTestType";

            StringAssert.DoesNotContain("mFooBar As MyTestType", refactoredCode.Source);
            StringAssert.Contains(destinationDeclaration, refactoredCode.Destination);
            StringAssert.Contains($"{moveDefinition.DestinationModuleName}.mFooBar.Foo", refactoredCode.Source);
            StringAssert.Contains($"{moveDefinition.DestinationModuleName}.mFooBar.Bar", refactoredCode.Source);
            if (moveDefinition.IsStdModuleSource && accessibility.Equals(Tokens.Public))
            {
                StringAssert.Contains($"{moveDefinition.DestinationModuleName}.{moveDefinition.SelectedElement}", refactoredCode[callSiteModuleName]);
                StringAssert.Contains($"result = ({moveDefinition.DestinationModuleName}.mFooBar.Foo + arg1) * 2", refactoredCode[callSiteModuleName]);
                StringAssert.Contains($"With {moveDefinition.DestinationModuleName}.mFooBar", refactoredCode[callSiteModuleName]);
                StringAssert.Contains($"NonQualifiedFoo = ({moveDefinition.DestinationModuleName}.mFooBar.Foo + arg1) * 2", refactoredCode[callSiteModuleName]);
            }
        }

        [TestCase("FormSource", ComponentType.UserForm)]
        [TestCase("ClassSource", ComponentType.ClassModule)]
        [TestCase("ModuleSource", ComponentType.StandardModule)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void XToStd_PrivateUDT_MoveField_NoStrategyFound(string moduleName, ComponentType componentType)
        {
            var source =
$@"
Option Explicit

Private Type MyTestType
    Foo As Long
    Bar As String
End Type

Private mFooBar As MyTestType

Public Function FooMath(arg1 As Long) As Long
    FooMath = arg1 * mFooBar.Foo
End Function

Public Function ConcatBar(arg1 As String) As String
    ConcatBar = arg1 & mFooBar.Bar
End Function";

            var resultCount = MoveMemberTestSupport.ParseAndTest(ThisTest, (moduleName, source, componentType));
            Assert.AreEqual(0, resultCount);

            long ThisTest(RubberduckParserState state, IVBE vbe, IRewritingManager rewritingManager)
            {
                var strategies = MoveMemberTestSupport.RetrieveStrategies(state, "mFooBar", DeclarationType.Variable, rewritingManager);
                return strategies.Count();
            }
        }

        [TestCase("Public")]
        [TestCase("Private")]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void StdToStd_PublicEnum_MoveField_HasReferences(string accessibility)
        {
            var source =
$@"
Option Explicit

Public Enum MyTestEnum
    ValueOne
    ValueTwo
End Enum

{accessibility} eFoo As MyTestEnum

Public Function FooMath(arg1 As Long) As Long
    FooMath = arg1 * eFoo
End Function
";

            var moveDefinition = new TestMoveDefinition(MoveEndpoints.StdToStd, ("eFoo", DeclarationType.Variable), sourceContent: source);

            var callSiteModuleName = "CallSiteModule";
            if (moveDefinition.IsStdModuleSource && accessibility.Equals(Tokens.Public))
            {
                var otherModuleReference =
    $@"
Option Explicit

Public Function MemberAccessFoo(arg1 As Long) As Long
    MemberAccessFoo = arg1 + {moveDefinition.SourceModuleName}.eFoo * 2
End Function

Public Function WithMemberAccessFoo(arg1 As Long) As Long
    Dim result As Long
    With {moveDefinition.SourceModuleName}
        result = (.eFoo + arg1) * 2
    End With
    WithMemberAccessFoo = result
End Function

Public Function NonQualifiedFoo(arg1 As Long) As Long
    NonQualifiedFoo = (eFoo + arg1) * 2
End Function
";
                moveDefinition.Add(new ModuleDefinition(callSiteModuleName, ComponentType.StandardModule, otherModuleReference));
            }

            var refactoredCode = RefactoredCode(moveDefinition, source);

            var destinationDeclaration = "Public eFoo As MyTestEnum";

            StringAssert.DoesNotContain("eFoo As MyTestEnum", refactoredCode.Source);
            StringAssert.Contains(destinationDeclaration, refactoredCode.Destination);
            StringAssert.Contains($"{moveDefinition.DestinationModuleName}.eFoo", refactoredCode.Source);
            if (moveDefinition.IsStdModuleSource && accessibility.Equals(Tokens.Public))
            {
                StringAssert.Contains($"{moveDefinition.DestinationModuleName}.{moveDefinition.SelectedElement}", refactoredCode[callSiteModuleName]);
                StringAssert.Contains($"result = ({moveDefinition.DestinationModuleName}.eFoo + arg1) * 2", refactoredCode[callSiteModuleName]);
                StringAssert.Contains($"NonQualifiedFoo = ({moveDefinition.DestinationModuleName}.eFoo + arg1) * 2", refactoredCode[callSiteModuleName]);
            }
        }

        [TestCase("FormSource", ComponentType.UserForm)]
        [TestCase("ClassSource", ComponentType.ClassModule)]
        [TestCase("ModuleSource", ComponentType.StandardModule)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void XToStd_PrivateEnum_MoveField_NoStrategyFound(string moduleName, ComponentType componentType)
        {
            var source =
$@"
Option Explicit

Private Enum MyTestEnum
    ValueOne
    ValueTwo
End Enum

Private eFoo As MyTestEnum

Public Function FooMath(arg1 As Long) As Long
    FooMath = arg1 * eFoo
End Function
";
            var resultCount = MoveMemberTestSupport.ParseAndTest(ThisTest, (moduleName, source, componentType));
            Assert.AreEqual(0, resultCount);

            long ThisTest(RubberduckParserState state, IVBE vbe, IRewritingManager rewritingManager)
            {
                var strategies = MoveMemberTestSupport.RetrieveStrategies(state, "eFoo", DeclarationType.Variable, rewritingManager);
                return strategies.Count();
            }
        }

        [TestCase("Public")]
        [TestCase("Private")]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void StdToStd_PublicArray_MoveField_HasReferences(string accessibility)
        {
            var source =
$@"
Option Explicit

{accessibility} mArray(5) As Long

Public Function FooMath(arg1 As Long) As Long
    FooMath = arg1 * mArray(2)
End Function
";

            var moveDefinition = new TestMoveDefinition(MoveEndpoints.StdToStd, ("mArray", DeclarationType.Variable), sourceContent: source);

            var callSiteModuleName = "CallSiteModule";
            if (moveDefinition.IsStdModuleSource && accessibility.Equals(Tokens.Public))
            {
                var otherModuleReference =
    $@"
Option Explicit

Public Function MemberAccessFoo(arg1 As Long) As Long
    MemberAccessFoo = arg1 + {moveDefinition.SourceModuleName}.mArray(3) * 2
End Function

Public Function WithMemberAccessFoo(arg1 As Long) As Long
    Dim result As Long
    With {moveDefinition.SourceModuleName}
        result = (.mArray(2) + arg1) * 2
    End With
    WithMemberAccessFoo = result
End Function

Public Function NonQualifiedFoo(arg1 As Long) As Long
    NonQualifiedFoo = (mArray(1) + arg1) * 2
End Function
";
                moveDefinition.Add(new ModuleDefinition(callSiteModuleName, ComponentType.StandardModule, otherModuleReference));
            }

            var refactoredCode = RefactoredCode(moveDefinition, source);

            var destinationDeclaration = "Public mArray(5) As Long";

            StringAssert.DoesNotContain("mArray(5) As Long", refactoredCode.Source);
            StringAssert.Contains(destinationDeclaration, refactoredCode.Destination);
            StringAssert.Contains($"{moveDefinition.DestinationModuleName}.mArray(2)", refactoredCode.Source);
            if (moveDefinition.IsStdModuleSource && accessibility.Equals(Tokens.Public))
            {
                StringAssert.Contains($"{moveDefinition.DestinationModuleName}.{moveDefinition.SelectedElement}(3)", refactoredCode[callSiteModuleName]);
                StringAssert.Contains($"result = ({moveDefinition.DestinationModuleName}.mArray(2) + arg1) * 2", refactoredCode[callSiteModuleName]);
                StringAssert.Contains($"NonQualifiedFoo = ({moveDefinition.DestinationModuleName}.mArray(1) + arg1) * 2", refactoredCode[callSiteModuleName]);
            }
        }

        //private IEnumerable<IMoveMemberRefactoringStrategy> RetrieveStrategies(IDeclarationFinderProvider declarationFinderProvider, string declarationName, DeclarationType declarationType, IRewritingManager rewritingManager)
        //{
        //    var target = declarationFinderProvider.DeclarationFinder.DeclarationsWithType(declarationType)
        //         .Single(declaration => declaration.IdentifierName == declarationName);

        //    var scenario = MoveMemberModel.CreateMoveScenario(declarationFinderProvider, target, new MoveDefinitionEndpoint("DefaultDestinationModule", ComponentType.StandardModule));
        //    var manager = new MoveMemberRewritingManager(rewritingManager);
        //    return MoveMemberStrategyProvider.FindStrategies(scenario, manager);
        //}
    }
}
