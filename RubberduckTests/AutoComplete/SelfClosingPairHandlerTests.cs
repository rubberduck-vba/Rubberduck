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
    public class SelfClosingPairTestInfo
    {
        public SelfClosingPairTestInfo(CodeString original, char input, CodeString rePrettified)
            : this(new Mock<SelfClosingPairCompletionService>(), original, original, input, rePrettified) { }

        public SelfClosingPairTestInfo(CodeString original, char input)
            : this(new Mock<SelfClosingPairCompletionService>(), original, original, input, original) { }

        public SelfClosingPairTestInfo(CodeString original, CodeString prettified, char input) 
            : this(new Mock<SelfClosingPairCompletionService>(), original, prettified, input, prettified) { }

        public SelfClosingPairTestInfo(Mock<SelfClosingPairCompletionService> service, CodeString original, CodeString prettified, char input, CodeString rePrettified, bool isControlKeyDown = false, bool isDeleteKey = false)
        {
            Original = original;
            Prettified = prettified;
            Input = input;
            RePrettified = rePrettified;
            Settings = AutoCompleteSettings.AllEnabled;

            Service = service;
            Module = new Mock<ICodeModule>();
            Handler = new Mock<ICodePaneHandler>();
            Handler.Setup(e => e.GetCurrentLogicalLine(Module.Object)).Returns(original);
            Handler.SetupSequence(e => e.Prettify(Module.Object, It.IsAny<CodeString>()))
                .Returns(prettified)
                .Returns(rePrettified);

            Args = new AutoCompleteEventArgs(Module.Object, input, isControlKeyDown, isDeleteKey);
        }

        public Mock<ICodeModule> Module { get; set; }
        public Mock<SelfClosingPairCompletionService> Service { get; set; }
        public Mock<ICodePaneHandler> Handler { get; set; }
        public CodeString Original { get; set; }
        public CodeString Prettified { get; set; }
        public char Input { get; set; }
        public CodeString RePrettified { get; set; }
        public AutoCompleteEventArgs Args { get; set; }
        public AutoCompleteSettings Settings { get; set; }

        public TestCodeString Result { get; set; }
    }

    [TestFixture]
    public class SelfClosingPairHandlerTests
    {
        private bool Run(SelfClosingPairTestInfo info)
        {
            var sut = new SelfClosingPairHandler(info.Handler.Object, info.Service.Object);
            if (sut.Handle(info.Args, info.Settings, out var result))
            {
                info.Result = new TestCodeString(result);
                return true;
            }

            info.Result = null;
            return false;
        }

        [Test]
        public void GivenInvalidInput_ResultIsNull()
        {
            var input = 'A'; // note: not a self-closing pair opening or closing character, not a handled key (e.g. '\b').
            var original = "DoSomething |".ToCodeString();
            var info = new SelfClosingPairTestInfo(original, input);

            Assert.IsFalse(Run(info));
            Assert.IsNull(info.Result);
        }

        [Test]
        public void GivenValidInput_InvokesSCP()
        {
            var input = '"'; // note: not a self-closing pair opening or closing character, not a handled key (e.g. '\b').
            var original = "DoSomething |".ToCodeString();
            var rePrettified = @"DoSomething ""|""".ToCodeString();
            var info = new SelfClosingPairTestInfo(original, input, rePrettified);

            Assert.IsTrue(Run(info));
            Assert.IsNotNull(info.Result);
        }

        [Test]
        public void GivenOpeningParenthesisOnOtherwiseNonEmptyLine_ReturnsFalseAndSwallowsKeypress()
        {
            var input = '(';
            var original = "foo = DateSerial(Year|)".ToCodeString();
            var rePrettified = "foo = DateSerial(Year(|))".ToCodeString();
            var info = new SelfClosingPairTestInfo(original, input, rePrettified);

            Assert.IsFalse(Run(info));
            Assert.IsTrue(info.Args.Handled);
        }

        [Test]
        public void GivenOpeningParenthesisOnCallStatement_ReturnsFalseAndLetsKeypressThrough()
        {
            var input = '(';
            var original = "Call DoSomething|".ToCodeString();
            var rePrettified = "Call DoSomething|".ToCodeString();
            var info = new SelfClosingPairTestInfo(original, input, rePrettified);

            Assert.IsFalse(Run(info));
            Assert.IsFalse(info.Args.Handled);
        }
        [Test]
        public void GivenBackspaceOnMatchedPair_DeletesMatchingTokens()
        {
            var input = '\b';
            var original = "foo = DateSerial(Year(|))".ToCodeString();
            var rePrettified = "foo = DateSerial(Year|)".ToCodeString();
            var info = new SelfClosingPairTestInfo(original, input, rePrettified);

            Assert.IsTrue(Run(info));
            Assert.IsNotNull(info.Result);
        }
    }
}
