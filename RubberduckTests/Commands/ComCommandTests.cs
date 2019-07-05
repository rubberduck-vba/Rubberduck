using System;
using Moq;
using Moq.Protected;
using NUnit.Framework;
using Rubberduck.UI.Command.ComCommands;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.Commands
{
    [TestFixture]
    [Category("ComCommand")]
    public class ComCommandTests
    {
        [Test]
        public void Verify_CanExecute_Before_And_After_Termination()
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(string.Empty, out _);
            var (command, vbeEvents) = ArrangeComCommand(vbe);
            vbeEvents.SetupSequence(v => v.Terminated).Returns(false).Returns(true);

            Assert.IsTrue(command.CanExecute(null));
            Assert.IsFalse(command.CanExecute(null));
        }

        [Test]
        public void Verify_OnExecute_Before_And_After_Termination()
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(string.Empty, out _);
            var (command, vbeEvents) = ArrangeComCommand(vbe);

            vbeEvents.SetupSequence(v => v.Terminated).Returns(false).Returns(true);
            command.Execute(null);
            command.Execute(null);

            command.VerifyOnExecute(Times.Once());
        }

        [Test]
        public void Verify_NoExecution_Terminated_BeforeCreation()
        {
            var vbeEvents = ArrangeVbeEvents();
            
            VbeEvents.Terminate();
            var command = ArranageComCommand(vbeEvents);
            command.Execute(null);

            command.VerifyOnExecute(Times.Never());
        }

        [Test]
        public void Verify_Execution_Among_Instances()
        {
            var vbeEvents = ArrangeVbeEvents();

            var command1 = ArranageComCommand(vbeEvents);
            command1.Execute(null);
            VbeEvents.Terminate();
            var command2 = ArranageComCommand(vbeEvents);
            command2.Execute(null);

            command1.VerifyOnExecute(Times.Once());
            command2.VerifyOnExecute(Times.Never());
        }

        [Test]
        public void Verify_Exception_Thrown_On_Null()
        {
            var vbe = new Mock<IVBE>();
            
            Assert.That(() =>
            {
                var vbeEvents = VbeEvents.Initialize(vbe.Object);
            }, Throws.TypeOf<NullReferenceException>());
        }

        private static VbeEvents ArrangeVbeEvents()
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule("", "foo", out _);
            return VbeEvents.Initialize(vbe.Object);
        }
        
        private static (ComCommandBase comCommand, Mock<IVbeEvents> vbeEvents) ArrangeComCommand(Mock<IVBE> vbe)
        {
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);
            return (ArranageComCommand(vbeEvents.Object), vbeEvents);
        }

        private static ComCommandBase ArranageComCommand(IVbeEvents vbeEvents)
        {
            // The ComCommandBase is an abstract class and is the subject under the test
            // Therefore, we actually want to use Moq.Mock to create an implementation
            // to directly test the base class' behaviors. We should not modify the mock
            // behavior, hence why we return the object, rather than the mock. 
            var mockComCommand = new Mock<ComCommandBase>(args:vbeEvents)
            {
                CallBase = true
            };
            return mockComCommand.Object;
        }
    }

    internal static class ComCommandExtensions
    {
        internal static void VerifyOnExecute(this ComCommandBase comCommand, Times times)
        {
            var mock = Mock.Get(comCommand);
            mock.Protected().Verify("OnExecute", times, ItExpr.IsAny<object>());
        }
    }
}

