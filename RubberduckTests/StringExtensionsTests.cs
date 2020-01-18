using System;
using NUnit.Framework;
using Rubberduck.Parsing.VBA.Extensions;
using Rubberduck.VBEditor.Extensions;

namespace RubberduckTests
{

    [TestFixture]
    public class StringExtensionsTests
    {
        [Test]
        [Category("Refactorings")]
        public void StripsStringLiteral()
        {
            var value = "\"Hello, World!\"";
            var instruction = "Debug.Print " + value;

            var result = instruction.StripStringLiterals();

            var replacement = new string(' ', value.Length);
            Assert.AreEqual("Debug.Print " + replacement, result);
        }

        [Test]
        [Category("Refactorings")]
        public void StripsAllStringLiterals()
        {
            var value = "\"Hello, World!\"";
            var instruction = "Debug.Print " + value + " & " + value;

            var result = instruction.StripStringLiterals();

            var replacement = new string(' ', value.Length);
            Assert.AreEqual("Debug.Print " + replacement + " & " + replacement, result);
        }

        [Test]
        [Category("Refactorings")]
        public void IsComment_StartLineWithSingleQuoteMarker()
        {
            var instruction = "'Debug.Print mwahaha this is just a comment.";

            int index;
            var result = instruction.HasComment(out index);

            Assert.IsTrue(result);
            Assert.AreEqual(index, 0);
        }

        [Test]
        [Category("Refactorings")]
        public void HasComment_EndOfLineSingleQuoteMarkerWithStringLiteral()
        {
            var comment = "'but this is one.";
            var instruction = "Debug.Print \"'this isn't a comment\" " + comment;

            int index;
            var result = instruction.HasComment(out index);

            Assert.IsTrue(result);
            Assert.AreEqual(comment, instruction.Substring(index));
        }

        [Test]
        [Category("Refactorings")]
        public void HasComment_RemMarkerWithWhitespace()
        {
            var comment = "Rem this is a comment.";
            var instruction = "Debug.Print \"'this isn't a comment\" : " + comment;

            int index;
            var result = instruction.HasComment(out index);

            Assert.IsTrue(result);
            Assert.AreEqual(comment, instruction.Substring(index));
        }

        [Test]
        [Category("Refactorings")]
        public void HasComment_RemMarkerWithQuestionMark()
        {
            var comment = "Rem?this is a comment.";
            var instruction = comment;

            int index;
            var result = instruction.HasComment(out index);

            Assert.IsTrue(result);
            Assert.AreEqual(comment, instruction.Substring(index));
        }

        [Test]
        [Category("Refactorings")]
        public void CaseInsensitiveContainsShouldReturnTrue()
        {
            var searchFor = "tExt";
            var textToSearch = "I contain some Text in here.";
            Assert.IsTrue(textToSearch.Contains(searchFor, StringComparison.OrdinalIgnoreCase));
        }

        [Test]
        [Category("Refactorings")]
        public void CaseInsensitiveContainsShouldReturnFalse()
        {
            var searchFor = "tExt";
            var textToSearch = "I don't have it.";
            Assert.IsFalse(textToSearch.Contains(searchFor, StringComparison.OrdinalIgnoreCase));
        }
    }
}
