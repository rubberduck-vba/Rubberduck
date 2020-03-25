using NUnit.Framework;
using Rubberduck.AutoComplete.SelfClosingPairs;
using Rubberduck.VBEditor;

namespace RubberduckTests.AutoComplete
{
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
        public void GivenDisabledSelfClosingPairs_BailsOut()
        {
            var input = '"';
            var original = "DoSomething |".ToCodeString();
            var info = new SelfClosingPairTestInfo(original, input);
            info.Settings.SelfClosingPairs.IsEnabled = false;

            Assert.IsFalse(Run(info));
            Assert.IsNull(info.Result);
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
        public void GivenOpeningParenthesisOnOtherwiseNonEmptyLine_ReturnsTrueAndSwallowsKeypress()
        {
            var input = '(';
            var original = "foo = DateSerial(Year|)".ToCodeString();
            var rePrettified = "foo = DateSerial(Year(|))".ToCodeString();
            var info = new SelfClosingPairTestInfo(original, input, rePrettified);

            Assert.IsTrue(Run(info));
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
        public void GivenOpeningCharInsideMultilineArgumentList_ReturnsTrueAndSwallowsKeypress()
        {
            var input = '"';
            var original = @"Err.Raise 5, _
     |".ToCodeString();
            var prettified = @"Err.Raise 5, _
|".ToCodeString();
            var rePrettified = @"Err.Raise 5, _
    ""|""".ToCodeString();
            var info = new SelfClosingPairTestInfo(original, prettified, input, rePrettified);

            Assert.IsTrue(Run(info));
            Assert.IsTrue(info.Args.Handled);
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
