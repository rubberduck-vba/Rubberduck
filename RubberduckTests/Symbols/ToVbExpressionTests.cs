using NUnit.Framework;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Extensions;

namespace RubberduckTests.Symbols
{
    [TestFixture]
    public class ToVbExpressionTests
    {
        [Test]
        [Category("String Extensions")]
        public void StringIsEnclosedInQuotes()
        {
            var managed = "foo";
            var expected = "\"foo\"";
            var actual = managed.ToVbExpression();

            Assert.AreEqual(expected, actual, "Expected {0}, actual was {1}", expected, actual);
        }

        [Test]
        [Category("String Extensions")]
        public void UseConstFlagOffGivesChrCall()
        {
            var managed = "\r";
            var expected = "Chr$(13)";
            var actual = managed.ToVbExpression(false);

            Assert.AreEqual(expected, actual, "Expected {0}, actual was {1}", expected, actual);
        }

        [Test]
        [Category("String Extensions")]
        public void UseConstFlagOnGivesConstant()
        {
            var managed = "\r";
            var expected = "vbCr";
            var actual = managed.ToVbExpression();

            Assert.AreEqual(expected, actual, "Expected {0}, actual was {1}", expected, actual);
        }

        [Test]
        [Category("String Extensions")]
        public void EmptyStringWorksConstFlagOff()
        {
            var managed = string.Empty;
            var expected = "\"\"";
            var actual = managed.ToVbExpression(false);

            Assert.AreEqual(expected, actual, "Expected {0}, actual was {1}", expected, actual);
        }

        [Test]
        [Category("String Extensions")]
        public void EmptyStringWorksConstFlagOn()
        {
            var managed = string.Empty;
            var expected = "vbNullString";
            var actual = managed.ToVbExpression();

            Assert.AreEqual(expected, actual, "Expected {0}, actual was {1}", expected, actual);
        }

        [Test]
        [Category("String Extensions")]
        public void NewLineIsASingleConstant()
        {
            var managed = "\r\n";
            var expected = "vbCrLf";
            var actual = managed.ToVbExpression();

            Assert.AreEqual(expected, actual, "Expected {0}, actual was {1}", expected, actual);
        }

        [Test]
        [Category("String Extensions")]
        public void MultipleNonPrintableCharactersWorksConstFlagOn()
        {
            var managed = "\r\x00";
            var expected = "vbCr & vbNullChar";
            var actual = managed.ToVbExpression();

            Assert.AreEqual(expected, actual, "Expected {0}, actual was {1}", expected, actual);
        }

        [Test]
        [Category("String Extensions")]
        public void MultipleNonPrintableCharactersWorksConstFlagOff()
        {
            var managed = "\r\x00";
            var expected = "Chr$(13) & Chr$(0)";
            var actual = managed.ToVbExpression(false);

            Assert.AreEqual(expected, actual, "Expected {0}, actual was {1}", expected, actual);
        }

        [Test]
        [Category("String Extensions")]
        public void MultilineStringWorksConstFlagOn()
        {
            var managed = "Line One\r\nLine Two";
            var expected = "\"Line One\" & vbCrLf & \"Line Two\"";
            var actual = managed.ToVbExpression();

            Assert.AreEqual(expected, actual, "Expected {0}, actual was {1}", expected, actual);
        }

        [Test]
        [Category("String Extensions")]
        public void MultilineStringWorksConstFlagOff()
        {
            var managed = "Line One\r\nLine Two";
            var expected = "\"Line One\" & Chr$(13) & Chr$(10) & \"Line Two\"";
            var actual = managed.ToVbExpression(false);

            Assert.AreEqual(expected, actual, "Expected {0}, actual was {1}", expected, actual);
        }

        [Test]
        [Category("String Extensions")]
        public void UnicodeUsesChrWCallWithHexNotation()
        {
            var managed = "•";
            var expected = "ChrW$(&H2022)";
            var actual = managed.ToVbExpression();

            Assert.AreEqual(expected, actual, "Expected {0}, actual was {1}", expected, actual);
        }

        [Test]
        [Category("String Extensions")]
        public void MixedAsciiAndUnicodeUsesChrAndChrWConstFlagOff()
        {
            var managed = "•\tBullet\x00";
            var expected = "ChrW$(&H2022) & Chr$(9) & \"Bullet\" & Chr$(0)";
            var actual = managed.ToVbExpression(false);

            Assert.AreEqual(expected, actual, "Expected {0}, actual was {1}", expected, actual);
        }

        [Test]
        [Category("String Extensions")]
        public void MixedAsciiAndUnicodeUsesChrAndChrWConstFlagOn()
        {
            var managed = "•\tBullet\x03";
            var expected = "ChrW$(&H2022) & vbTab & \"Bullet\" & Chr$(3)";
            var actual = managed.ToVbExpression();

            Assert.AreEqual(expected, actual, "Expected {0}, actual was {1}", expected, actual);
        }

        [Test]
        [Category("String Extensions")]
        public void NonBreakingSpaceIsReplaced()    //Looking at you, @ThunderFrame
        {
            var managed = "\xA0";
            var expected = "Chr$(160)";
            var actual = managed.ToVbExpression();

            Assert.AreEqual(expected, actual, "Expected {0}, actual was {1}", expected, actual);
        }
    }
}
