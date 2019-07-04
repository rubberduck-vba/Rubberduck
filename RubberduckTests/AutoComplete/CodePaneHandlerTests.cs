using Moq;
using NUnit.Framework;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SourceCodeHandling;
using RubberduckTests.Mocks;

namespace RubberduckTests.AutoComplete
{
    [TestFixture]
    public class CodePaneHandlerTests
    {
        [Test]
        [Category("AutoComplete")]
        public void ActuallyDeletesAndInsertsOriginalLine()
        {
            var code = "MsgBox|".ToCodeString();

            var sut = InitializeSut(code, code, out var module, out _);
            sut.Prettify(module.Object, code);

            module.Verify(m => m.DeleteLines(code.SnippetPosition.StartLine, code.SnippetPosition.LineCount), Times.Once);
            module.Verify(m => m.InsertLines(code.SnippetPosition.StartLine, code.Code), Times.Once);
        }

        [Test]
        [Category("AutoComplete")]
        public void GivenSamePrettifiedCode_YieldsSameCodeString()
        {
            var original = "MsgBox (|".ToCodeString();

            var sut = InitializeSut(original, original, out var module, out _);
            var actual = new TestCodeString(sut.Prettify(module.Object, original));

            Assert.AreEqual(original, actual);
        }

        [Test]
        [Category("AutoComplete")]
        public void GivenLeadingWhitespace_YieldsSameCodeString()
        {
            var original = "    MsgBox|".ToCodeString();

            var sut = InitializeSut(original, original, out var module, out _);
            var actual = new TestCodeString(sut.Prettify(module.Object, original));

            Assert.AreEqual(original, actual);
        }

        [Test]
        [Category("AutoComplete")]
        public void GivenTrailingWhitespace_IsTrimmedAndPrettifiedCaretIsAtLastCharacter()
        {
            var original = "MsgBox |".ToCodeString();
            var prettified = "MsgBox".ToCodeString();
            var expected = "MsgBox|".ToCodeString();

            var sut = InitializeSut(original, prettified, out var module, out _);
            var actual = new TestCodeString(sut.Prettify(module.Object, original));

            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("AutoComplete")]
        public void GivenPrintToken_PrettifiedCaretIsAtLastCharacter()
        {
            var original = "debug.? dosomething|".ToCodeString();
            var prettified = "Debug.Print DoSomething".ToCodeString();
            var expected = "Debug.Print DoSomething|".ToCodeString();

            var sut = InitializeSut(original, prettified, out var module, out _);
            var actual = new TestCodeString(sut.Prettify(module.Object, original));

            Assert.AreEqual(expected, actual);
        }

    [Test]
        [Category("AutoComplete")]
        public void GivenExtraWhitespace_PrettifiedCaretStillAtSameToken()
        {
            var original = "MsgBox      (\"test|\")".ToCodeString();
            var prettified = "MsgBox (\"test\")".ToCodeString();
            var expected = "MsgBox (\"test|\")".ToCodeString();

            var sut = InitializeSut(original, prettified, out var module, out _);
            var actual = new TestCodeString(sut.Prettify(module.Object, original));
            
            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("AutoComplete")]
        public void GivenMultilineLogicalLine_TracksCaret()
        {
            var original = @"
MsgBox ""test"" & vbNewLine & _
       ""|"")".ToCodeString();

            var sut = InitializeSut(original, original, out var module, out _);
            var actual = new TestCodeString(sut.Prettify(module.Object, original));

            Assert.AreEqual(original, actual);
        }

        [Test]
        [Category("AutoComplete")]
        public void GivenPartialMultilineInstruction_TracksCaret()
        {
            var original = @"
ExecuteStoredProcedure (""AddAppointmentCountForAClinic"", False,dbconfig.SQLConString, _
                | thisClinic.ClinicID ,".ToCodeString();

            var sut = InitializeSut(original, original, out var module, out _);
            var actual = new TestCodeString(sut.Prettify(module.Object, original));

            Assert.AreEqual(original, actual);
        }

        private static ICodePaneHandler InitializeSut(TestCodeString original, TestCodeString prettified, out Mock<ICodeModule> module, out Mock<ICodePane> pane)
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

            var sut = new CodePaneHandler(new ProjectsRepository(vbe.Object));
            return sut;
        }
    }
}