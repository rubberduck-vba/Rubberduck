using NUnit.Framework;
using Moq;
using Rubberduck.AutoComplete.Service;
using Rubberduck.Settings;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SourceCodeHandling;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace RubberduckTests.AutoComplete
{
    [TestFixture]
    public class SelfClosingPairHandlerTests
    {
        private bool Run(CodeString original, CodeString prettified, char input, CodeString rePrettified, out TestCodeString testResult, bool isControlKeyDown = false, bool isDeleteKey = false)
        {
            var service = new Mock<SelfClosingPairCompletionService>();
            return Run(service, original, prettified, input, rePrettified, out testResult, isControlKeyDown, isDeleteKey);
        }

        private bool Run(Mock<SelfClosingPairCompletionService> service, CodeString original, CodeString prettified, char input,  CodeString rePrettified, out TestCodeString testResult, bool isControlKeyDown = false, bool isDeleteKey = false)
        {
            var module = new Mock<ICodeModule>();
            var handler = new Mock<ICodePaneHandler>();
            handler.Setup(e => e.GetCurrentLogicalLine(module.Object)).Returns(original);
            handler.SetupSequence(e => e.Prettify(module.Object, It.IsAny<CodeString>()))
                .Returns(prettified)
                .Returns(rePrettified);

            var settings = AutoCompleteSettings.AllEnabled;

            var args = new AutoCompleteEventArgs(module.Object, input, isControlKeyDown, isDeleteKey);
            var sut = new SelfClosingPairHandler(handler.Object, service.Object);

            if (sut.Handle(args, settings, out var result))
            {
                testResult = new TestCodeString(result);
                return true;
            }

            testResult = null;
            return false;
        }

        [Test]
        public void GivenInvalidInput_ResultIsNull()
        {
            var input = 'A'; // note: not a self-closing pair opening or closing character, not a handled key (e.g. '\b').
            var original = "DoSomething |".ToCodeString();

            Assert.IsFalse(Run(original, original, input, original, out var result));
            Assert.IsNull(result);
        }

        [Test]
        public void GivenValidInput_InvokesSCP()
        {
            var input = '"'; // note: not a self-closing pair opening or closing character, not a handled key (e.g. '\b').
            var original = "DoSomething |".ToCodeString();
            var rePrettified = @"DoSomething ""|""".ToCodeString();

            Assert.IsTrue(Run(original, original, input, rePrettified, out var result));
            Assert.IsNotNull(result);
        }

        [Test]
        public void GivenOpeningParenthesisOnOtherwiseNonEmptyLine_ReturnsFalse()
        {
            var input = '(';
            var original = "foo = DateSerial(Year|)".ToCodeString();
            var rePrettified = "foo = DateSerial(Year(|))".ToCodeString();

            Assert.IsFalse(Run(original, original, input, rePrettified, out var result));
            Assert.IsNull(result);
        }
    }
}
