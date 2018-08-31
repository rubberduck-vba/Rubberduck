using NUnit.Framework;
using Rubberduck.VBEditor.SafeComWrappers;
using System;

namespace RubberduckTests.Symbols
{
    [TestFixture]
    public class CommandBarControlCaptionGuardTests
    {
        private const int MAX_CAPTION_LENGTH = 256;
        private const string COORDINATE = "L5C27";
        private const string MODULENAME = "TheContainingModule";
        private const string FILENAME = "TheFilename";
        private const string TYPEIDENTIFIER = "(procedure)";

        [Test]
        [Category("Guard")]
        public void CommandBarControlCaptionGuard_Regex_MeetsPattern()
        {
            string proposedCaption = GetFormattedMethodIdentifierOfLength(60);
            Assert.IsTrue(CommandBarControlCaptionGuard.IsMethodFormat(proposedCaption));
        }

        [Test]
        [Category("Guard")]
        public void CommandBarControlCaptionGuard_Regex_TooShort()
        {
            string proposedCaption = GetFormattedMethodIdentifierOfLength(60);
            Assert.IsFalse(CommandBarControlCaptionGuard.IsMethodFormat(proposedCaption.Substring(0, 5)));
        }

        [Test]
        [Category("Guard")]
        public void CommandBarControlCaptionGuard_Regex_ExtraChars()
        {
            string proposedCaption = GetFormattedMethodIdentifierOfLength(60);
            Assert.IsFalse(CommandBarControlCaptionGuard.IsMethodFormat(proposedCaption + "Yo!!"));
        }

    [Test]
        [Category("GuardFunction")]
        public void CommandBarControlCaptionGuard_ShortName()
        {
            string proposedCaption = GetFormattedMethodIdentifierOfLength(60);
            var result = CommandBarControlCaptionGuard.ApplyGuard(proposedCaption);

            Assert.IsTrue(result.Equals(proposedCaption, StringComparison.InvariantCulture));
            Assert.IsFalse(result.Contains("..."));
        }

        [Test]
        [Category("GuardFunction")]
        public void CommandBarControlCaptionGuard_TooLongSubName()
        {
            string proposedCaption = GetFormattedMethodIdentifierOfLength(260);
            var result = CommandBarControlCaptionGuard.ApplyGuard(proposedCaption);

            Assert.IsTrue(result.Length <= MAX_CAPTION_LENGTH);
            Assert.IsTrue(result.Contains("..."));
        }

        [Test]
        [Category("GuardFunction")]
        public void CommandBarControlCaptionGuard_TooLongFileName()
        {
            string reallyLongFilename = GetIdentifierOfLength(200);
            string proposedCaption = GetFormattedMethodIdentifierOfLength(300, reallyLongFilename);
            var result = CommandBarControlCaptionGuard.ApplyGuard(proposedCaption);

            Assert.IsTrue(result.Length <= MAX_CAPTION_LENGTH);
            Assert.IsTrue(result.Contains("..."));
        }

        private string GetFormattedMethodIdentifierOfLength(int targetLength, string fileName = FILENAME)
        {
            var preamble = $"{COORDINATE} | {fileName}.{MODULENAME}.";
            var proposedCaption = preamble + GetIdentifierOfLength(targetLength - preamble.Length);
            proposedCaption = proposedCaption.Substring(0, targetLength - (TYPEIDENTIFIER.Length + 1)) + " " + TYPEIDENTIFIER;
            if (proposedCaption.Length != targetLength)
            {
                Assert.Inconclusive("Test Generated Format String has incorrect length");
            }
            return proposedCaption;
        }

        private string GetIdentifierOfLength(int targetLength)
        {
            int maxLoopIterations = 100;
            string proposedCaption = "Abcdefghij";
            for (int idx = 0; idx <= maxLoopIterations; idx++)
            {
                proposedCaption = proposedCaption + proposedCaption;
                if (proposedCaption.Length > targetLength)
                {
                    idx = maxLoopIterations + 1;
                    proposedCaption = proposedCaption.Substring(0, targetLength);
                }
            }
            if (proposedCaption.Length != targetLength)
            {
                Assert.Inconclusive("Test Generated String has incorrect length");
            }
            return proposedCaption;
        }
    }
}
