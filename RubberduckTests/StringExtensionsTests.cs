using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBA;
using Rubberduck.VBEditor.Extensions;

namespace RubberduckTests
{
    [TestClass]
    public class StringExtensionsTests
    {
        [TestMethod, Timeout(1000)]
        public void StripsStringLiteral()
        {
            var value = "\"Hello, World!\"";
            var instruction = "Debug.Print " + value;

            var result = instruction.StripStringLiterals();

            var replacement = new string(' ', value.Length);
            Assert.AreEqual("Debug.Print " + replacement, result);
        }

        [TestMethod, Timeout(1000)]
        public void StripsAllStringLiterals()
        {
            var value = "\"Hello, World!\"";
            var instruction = "Debug.Print " + value + " & " + value;

            var result = instruction.StripStringLiterals();

            var replacement = new string(' ', value.Length);
            Assert.AreEqual("Debug.Print " + replacement + " & " + replacement, result);
        }

        [TestMethod, Timeout(1000)]
        public void IsComment()
        {
            var instruction = "'Debug.Print mwahaha this is just a comment.";

            int index;
            var result = instruction.HasComment(out index);

            Assert.IsTrue(result);
            Assert.AreEqual(index, 0);
        }

        [TestMethod, Timeout(1000)]
        public void HasComment()
        {
            var comment = "'but this is one.";
            var instruction = "Debug.Print \"'this isn't a comment\" " + comment;

            int index;
            var result = instruction.HasComment(out index);

            Assert.IsTrue(result);
            Assert.AreEqual(comment, instruction.Substring(index));
        }

        [TestMethod, Timeout(1000)]
        public void CaseInsensitiveContainsShouldReturnTrue()
        {
            var searchFor = "tExt";
            var textToSearch = "I contain some Text in here.";
            Assert.IsTrue(textToSearch.Contains(searchFor, StringComparison.OrdinalIgnoreCase));
        }

        [TestMethod, Timeout(1000)]
        public void CaseInsensitiveContainsShouldReturnFalse()
        {
            var searchFor = "tExt";
            var textToSearch = "I don't have it.";
            Assert.IsFalse(textToSearch.Contains(searchFor, StringComparison.OrdinalIgnoreCase));
        }
    }
}
