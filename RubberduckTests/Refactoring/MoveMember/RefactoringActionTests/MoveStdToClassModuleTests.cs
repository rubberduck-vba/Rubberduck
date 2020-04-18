using NUnit.Framework;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.MoveMember;

namespace RubberduckTests.Refactoring.MoveMember
{
    [TestFixture]
    class MoveStdToClassModuleTests : MoveMemberRefactoringActionTestSupportBase
    {
        [Test]
        [Ignore("StdToClass - Future")]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void MovedPropertyCallsRetainedProperty()
        {
            var memberToMove = ("FooTimes2", DeclarationType.PropertyGet);
            var endpoints = MoveEndpoints.StdToClass;
            var source =
$@"
Option Explicit

Private Const mFoo As Long = 10

Public Property Get Foo() As Long
    Foo = mFoo
End Property

Public Property Get FooTimes2() As Long 
    FooTimes2 = Foo * 2
End Property
";

            MoveMemberModel modelAdjustment(MoveMemberModel model)
            {
                model.ChangeDestination(endpoints.DestinationModuleName(), endpoints.DestinationComponentType());
                return model;
            }

            var refactoredCode = RefactorTargets(memberToMove, endpoints, source, string.Empty, modelAdjustment);

            StringAssert.Contains("Get Foo()", refactoredCode.Source);
            StringAssert.Contains($"FooTimes2 = {endpoints.SourceModuleName()}.Foo * 2", refactoredCode.Destination);
            StringAssert.Contains("Private Const mFoo As Long = 10", refactoredCode.Source);
        }

        [Test]
        [Ignore("StdToClass - Future")]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void RetainedPropertyCallsMovedProperty()
        {
            var memberToMove = ("Foo", DeclarationType.PropertyGet);
            var endpoints = MoveEndpoints.StdToClass;
            var source =
$@"
Option Explicit

Private Const mFoo As Long = 10

Public Property Get Foo() As Long
    Foo = mFoo
End Property

Public Property Get FooTimes2() As Long 
    FooTimes2 = Foo * 2
End Property
";

            MoveMemberModel modelAdjustment(MoveMemberModel model)
            {
                model.ChangeDestination(endpoints.DestinationModuleName(), endpoints.DestinationComponentType());
                return model;
            }

            var refactoredCode = RefactorTargets(memberToMove, endpoints, source, string.Empty, modelAdjustment);


            StringAssert.Contains($"Private {endpoints.DestinationClassInstanceName()} As {endpoints.DestinationModuleName()}", refactoredCode.Source);
            StringAssert.Contains($"Set {endpoints.DestinationClassInstanceName()} = New {endpoints.DestinationModuleName()}", refactoredCode.Source);
            StringAssert.Contains($"FooTimes2 = {endpoints.DestinationModuleName()}.Foo * 2", refactoredCode.Source);

            StringAssert.Contains("Get Foo()", refactoredCode.Destination);
            StringAssert.Contains("Private Const mFoo As Long = 10", refactoredCode.Destination);
        }

        [Test]
        [Ignore("StdToClass - Future")]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void UsesExistingObjectFieldInstance()
        {
            var memberToMove = ("FooTimes2", DeclarationType.PropertyGet);
            var endpoints = MoveEndpoints.StdToClass;
            var source =
$@"
Option Explicit

Private {endpoints.DestinationClassInstanceName()} As {endpoints.DestinationModuleName()}

Private Property Get {endpoints.DestinationModuleName()}() As {endpoints.DestinationModuleName()}
    Set {endpoints.DestinationModuleName()} = mFoo
End Property

Public Property Get FooTimes2() As Long 
    FooTimes2 = {endpoints.DestinationModuleName()}.Foo * 2
End Property

Public Property Get FooTimes4() As Long 
    FooTimes2 = {endpoints.DestinationModuleName()}.Foo * 4
End Property
";

            var destination =
$@"
Option Explicit

Private Const mFoo As Long = 10

Public Property Get Foo() As Long
    Foo = mFoo
End Property
";

            MoveMemberModel modelAdjustment(MoveMemberModel model)
            {
                model.ChangeDestination(endpoints.DestinationModuleName(), endpoints.DestinationComponentType());
                return model;
            }

            var refactoredCode = RefactorTargets(memberToMove, endpoints, source, string.Empty, modelAdjustment);


            Assert.IsTrue($"Private {endpoints.DestinationClassInstanceName()} As {endpoints.DestinationModuleName()}".OccursOnce(refactoredCode.Source));
            StringAssert.Contains($"Set {endpoints.DestinationClassInstanceName()} = New {endpoints.DestinationModuleName()}", refactoredCode.Source);
            StringAssert.Contains($"FooTimes4 = {endpoints.DestinationModuleName()}.Foo * 4", refactoredCode.Source);

            StringAssert.Contains("Get Foo()", refactoredCode.Destination);
            StringAssert.Contains("Private Const mFoo As Long = 10", refactoredCode.Destination);
        }
    }
}
