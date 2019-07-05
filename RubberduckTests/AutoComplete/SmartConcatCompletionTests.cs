using System.Linq;
using Moq;
using NUnit.Framework;
using Rubberduck.AutoComplete.SmartConcat;
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
    [TestFixture]
    public class SmartConcatCompletionTests
    {
        [Test]
        public void MaintainsIndent()
        {
            var original = "foo = \"a|\"".ToCodeString();
            var expected = original.Lines[0].IndexOf('"');

            var result = Run(original, '\r');
            if (result.Lines.Length != original.Lines.Length + 1)
            {
                Assert.Inconclusive();
            }
            var actual = result.Lines[1].IndexOf('"');

            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void CtrlEnterAddsVbNewLineToken()
        {
            var original = "foo = \"a|\"".ToCodeString();
            var expected = "foo = \"a\" & vbNewLine & _";

            var result = Run(original, '\r', true);
            var actual = result.Lines[0];

            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void PlacesCaretOnNextLineBetweenStringDelimiters()
        {
            var original = "foo = \"a|\"".ToCodeString();
            var expected = "foo = \"a\" & _\r\n      \"|\"".ToCodeString();

            var actual = Run(original, '\r');
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void WorksGivenCaretOnSecondPhysicalCodeLine()
        {
            var original = "foo = \"a\" & _\r\n      \"|\"".ToCodeString();
            var expected = "foo = \"a\" & _\r\n      \"\" & _\r\n      \"|\"".ToCodeString();

            var actual = Run(original, '\r');
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void SplittingExistingString_PutsCaretAtSameRelativePosition()
        {
            var original = "foo = \"ab|cd\"".ToCodeString();
            var expected = "foo = \"ab\" & _\r\n      \"|cd\"".ToCodeString();

            var actual = Run(original, '\r');
            Assert.AreEqual(expected, actual);
        }

        private static TestCodeString Run(TestCodeString original, char input, bool isCtrlDown = false, bool isDeleteKey = false)
        {

            var sut = InitializeSut(original, out var module, out var settings);
            var args = new AutoCompleteEventArgs(module.Object, input, isCtrlDown, isDeleteKey);

            if (sut.Handle(args, settings, out var result))
            {
                return new TestCodeString(result);
            }

            return null;
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
            var paneSelection = new Selection(original.SnippetPosition.StartLine + original.CaretPosition.StartLine, original.CaretPosition.StartColumn + 1);
            pane.Object.Selection = paneSelection;

            module.Setup(m => m.DeleteLines(original.SnippetPosition.StartLine, original.SnippetPosition.LineCount));
            module.Setup(m => m.InsertLines(original.SnippetPosition.StartLine, original.Code));
            module.Setup(m => m.CodePane).Returns(pane.Object);
            for (var i = 0; i < original.SnippetPosition.LineCount; i++)
            {
                var index = i;
                module.Setup(m => m.GetLines(index + 1, 1)).Returns(original.Lines[index]);
            }
            module.Setup(m => m.GetLines(original.SnippetPosition)).Returns(prettified.Code);
            module.Setup(m => m.GetLines(paneSelection.StartLine, paneSelection.LineCount)).Returns(prettified.CaretLine);

            settings = new AutoCompleteSettings {IsEnabled = true};
            settings.SmartConcat.IsEnabled = true;
            settings.SmartConcat.ConcatVbNewLineModifier = ModifierKeySetting.CtrlKey;
            settings.SmartConcat.ConcatMaxLines = AutoCompleteSettings.ConcatMaxLinesMaxValue;

            var handler = new CodePaneHandler(new ProjectsRepository(vbe.Object));
            var sut = new SmartConcatenationHandler(handler);
            return sut;
        }
    }
}