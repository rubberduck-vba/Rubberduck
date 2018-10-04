using System.Linq;
using Moq;
using NUnit.Framework;
using Rubberduck.AutoComplete;
using Rubberduck.AutoComplete.Service;
using Rubberduck.Settings;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SourceCodeHandling;
using RubberduckTests.Mocks;

namespace RubberduckTests.AutoComplete
{
    [TestFixture][Ignore("nothing to see here. yet.")]
    public class SmartConcatCompletionTests
    {
        [Test]
        public void MaintainsIndent()
        {
            var original = "foo = \"a|\"".ToCodeString();
            var expected = "foo = \"a\" & _\r\n      \"|\"".ToCodeString();

            var sut = InitializeSut(original, out var module, out var settings);
            var args = new AutoCompleteEventArgs(module.Object, '\r', false, false);

            sut.Handle(args, settings);
        }

        private static SmartConcatenationHandler InitializeSut(TestCodeString code, out Mock<ICodeModule> module, out AutoCompleteSettings settings)
        {
            return InitializeSut(code, code, out module, out _, out settings);
        }

        private static SmartConcatenationHandler InitializeSut(TestCodeString original, TestCodeString prettified, out Mock<ICodeModule> module, out Mock<ICodePane> pane, out AutoCompleteSettings settings)
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

            settings = new AutoCompleteSettings(Enumerable.Empty<AutoCompleteSetting>()) { EnableSmartConcat = true };

            var handler = new CodePaneSourceCodeHandler(new ProjectsRepository(vbe.Object));
            var sut = new SmartConcatenationHandler(handler);
            return sut;
        }
    }
}