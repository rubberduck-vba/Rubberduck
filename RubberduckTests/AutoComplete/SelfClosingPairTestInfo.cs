using Moq;
using Rubberduck.AutoComplete;
using Rubberduck.AutoComplete.SelfClosingPairs;
using Rubberduck.Settings;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SourceCodeHandling;

namespace RubberduckTests.AutoComplete
{
    public class SelfClosingPairTestInfo
    {
        public SelfClosingPairTestInfo(CodeString original, char input, CodeString rePrettified)
            : this(new Mock<SelfClosingPairCompletionService>(ArrangeShowQuickCommand().Object), original, original, input, rePrettified) { }
        public SelfClosingPairTestInfo(CodeString original, CodeString prettified, char input, CodeString rePrettified)
            : this(new Mock<SelfClosingPairCompletionService>(ArrangeShowQuickCommand().Object), original, prettified, input, rePrettified) { }

        public SelfClosingPairTestInfo(CodeString original, char input)
            : this(new Mock<SelfClosingPairCompletionService>(ArrangeShowQuickCommand().Object), original, original, input, original) { }

        public SelfClosingPairTestInfo(
            Mock<SelfClosingPairCompletionService> service, 
            CodeString original, 
            CodeString prettified, 
            char input,
            CodeString rePrettified, 
            bool isControlKeyDown = false, 
            bool isDeleteKey = false)
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

        public static Mock<IShowQuickInfoCommand> ArrangeShowQuickCommand()
        {
            return new Mock<IShowQuickInfoCommand>();
        }
    }
}