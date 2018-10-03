using Moq;
using NUnit.Framework;
using Rubberduck.AutoComplete;
using Rubberduck.AutoComplete.SelfClosingPairCompletion;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SourceCodeHandling;
using RubberduckTests.Mocks;

namespace RubberduckTests.AutoComplete
{
    [TestFixture][Ignore("need to extract class for this sut.")]
    public class SmartConcatCompletionTests
    {
        [Test]
        public void MaintainsIndent()
        {
            var original = "foo = \"a|\"".ToCodeString();
            var expected = "foo = \"a\" & _\r\n      \"|\"".ToCodeString();

            var sut = InitializeSut(original, out var module);
            var args = new AutoCompleteEventArgs(module.Object, '\r', false, false);

            sut.Run(args);
        }

        private static AutoCompleteKeyDownHandler InitializeSut(TestCodeString code, out Mock<ICodeModule> module)
        {
            return InitializeSut(code, code, out module, out _);
        }

        private static AutoCompleteKeyDownHandler InitializeSut(TestCodeString original, TestCodeString prettified, out Mock<ICodeModule> module, out Mock<ICodePane> pane)
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, "");
            var vbe = builder.AddProject(project.Build()).Build();

            module = new Mock<ICodeModule>();
            pane = new Mock<ICodePane>();
            pane.SetupProperty(m => m.Selection);
            module.Setup(m => m.DeleteLines(original.SnippetPosition.StartLine, original.SnippetPosition.LineCount));
            module.Setup(m => m.InsertLines(original.SnippetPosition.StartLine, original.Code));
            module.Setup(m => m.CodePane).Returns(pane.Object);
            module.Setup(m => m.GetLines(original.SnippetPosition)).Returns(prettified.Code);

            var handler = new CodePaneSourceCodeHandler(new ProjectsRepository(vbe.Object));
            var sut = new AutoCompleteKeyDownHandler(handler, new SelfClosingPairCompletionService(new Mock<IShowIntelliSenseCommand>().Object));

            return sut;
        }
    }
}