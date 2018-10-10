using System;
using Antlr4.Runtime;
using Antlr4.Runtime.Atn;
using Antlr4.Runtime.Tree;
using Moq;
using NUnit.Framework;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.Parsing.VBA.Parsing.ParsingExceptions;

namespace RubberduckTests.Parsing.Coordination
{
    [TestFixture]
    public class TokenStreamParserTests
    {
        [Category("Parser")]
        [TestCase(ParserMode.FallBackSllToLl, 1, 0)]
        [TestCase(ParserMode.SllOnly, 1, 0)]
        [TestCase(ParserMode.LlOnly, 0, 1)]
        public void TokenStreamParserCallsParserModesTheExpectedAmountOfTimesWhenThereIsNoException(ParserMode parseMode, int expectedSllCount, int expectedLlCount)
        {
            var errorListenerFactory = DummyErrorListenerFactory();
            var llCount = 0;
            var sllCount = 0;
            
            var llFunction = new Func<IParseTree>(() =>
            {
                llCount++;
                return null;
            });
            var sllFunction = new Func<IParseTree>(() =>
            {
                sllCount++;
                return null;
            });

            var tokenStreamParser = new TestTokenStreamParser(errorListenerFactory, errorListenerFactory, llFunction, sllFunction);
            tokenStreamParser.Parse("TestModule", null, CodeKind.CodePaneCode, parseMode);

            Assert.AreEqual(expectedLlCount, llCount);
            Assert.AreEqual(expectedSllCount, sllCount);
        }

        private IParsePassErrorListenerFactory DummyErrorListenerFactory()
        {
            Exception exception = null;
            var errorListenerMock = new Mock<IRubberduckParseErrorListener>();
            errorListenerMock.Setup(m => m.HasPostponedException(out exception))
                .Returns(() => false);
            var errorListenerFactoryMock = new Mock<IParsePassErrorListenerFactory>();
            errorListenerFactoryMock.Setup(m => m.Create(It.IsAny<string>(), It.IsAny<CodeKind>()))
                .Returns((string s, CodeKind ck) => errorListenerMock.Object);
            return errorListenerFactoryMock.Object;
        }

        [Category("Parser")]
        [Test]
        public void TokenStreamParserCallsLlModeIfSllFailsInFallBackMode()
        {
            var errorListenerFactory = DummyErrorListenerFactory();
            var llCount = 0;
            var sllCount = 0;

            var llFunction = new Func<IParseTree>(() =>
            {
                llCount++;
                return null;
            });
            var sllFunction = new Func<IParseTree>(() =>
            {
                sllCount++;
                throw new Exception();
            });

            var tokenStreamParser = new TestTokenStreamParser(errorListenerFactory, errorListenerFactory, llFunction, sllFunction);
            tokenStreamParser.Parse("TestModule", null, CodeKind.CodePaneCode, ParserMode.FallBackSllToLl);

            Assert.AreEqual(1, llCount);
            Assert.AreEqual(1, sllCount);
        }

        [Category("Parser")]
        [Test]
        public void TokenStreamParserThrowsForExceptionInSllForSllOnlyMode()
        {
            var errorListenerFactory = DummyErrorListenerFactory();

            var llFunction = new Func<IParseTree>(() => null);
            var sllFunction = new Func<IParseTree>(() => throw new TokenStreamParserTestException());

            var tokenStreamParser = new TestTokenStreamParser(errorListenerFactory, errorListenerFactory, llFunction, sllFunction);
            Assert.Throws<TokenStreamParserTestException>(() => tokenStreamParser.Parse("TestModule", null, CodeKind.CodePaneCode, ParserMode.SllOnly));
        }

        [Category("Parser")]
        [Test]
        public void TokenStreamParserThrowsForExceptionInLlForLlOnlyMode()
        {
            var errorListenerFactory = DummyErrorListenerFactory();

            var llFunction = new Func<IParseTree>(() => throw new TokenStreamParserTestException());
            var sllFunction = new Func<IParseTree>(() => null);

            var tokenStreamParser = new TestTokenStreamParser(errorListenerFactory, errorListenerFactory, llFunction, sllFunction);
            Assert.Throws<TokenStreamParserTestException>(() => tokenStreamParser.Parse("TestModule", null, CodeKind.CodePaneCode, ParserMode.LlOnly));
        }

        [Category("Parser")]
        [Test]
        public void TokenStreamParserThrowsForExceptionInBothModesForFallbackMode()
        {
            var errorListenerFactory = DummyErrorListenerFactory();

            var llFunction = new Func<IParseTree>(() => throw new TokenStreamParserTestException());
            var sllFunction = new Func<IParseTree>(() => throw new Exception());

            var tokenStreamParser = new TestTokenStreamParser(errorListenerFactory, errorListenerFactory, llFunction, sllFunction);
            Assert.Throws<TokenStreamParserTestException>(() => tokenStreamParser.Parse("TestModule", null, CodeKind.CodePaneCode, ParserMode.FallBackSllToLl));
        }

        [Category("Parser")]
        [TestCase(ParserMode.SllOnly)]
        [TestCase(ParserMode.LlOnly)]
        [TestCase(ParserMode.FallBackSllToLl)]
        public void TokenStreamParserThrowsPostponedExceptionAfterFinishingWithoutNonPostponedExceptions(ParserMode parseMode)
        {
            var exception = new TokenStreamParserTestException();
            var errorListenerFactory = ErrorListenerFactoryWithPostponedException(exception);

            var llFunction = new Func<IParseTree>(() => null);
            var sllFunction = new Func<IParseTree>(() => null);

            var tokenStreamParser = new TestTokenStreamParser(errorListenerFactory, errorListenerFactory, llFunction, sllFunction);
            Assert.Throws<TokenStreamParserTestException>(() => tokenStreamParser.Parse("TestModule", null, CodeKind.CodePaneCode, ParserMode.FallBackSllToLl));
        }

        private IParsePassErrorListenerFactory ErrorListenerFactoryWithPostponedException(Exception exception)
        {
            var errorListenerMock = new Mock<IRubberduckParseErrorListener>();
            errorListenerMock.Setup(m => m.HasPostponedException(out exception))
                .Returns(() => true);
            var errorListenerFactoryMock = new Mock<IParsePassErrorListenerFactory>();
            errorListenerFactoryMock.Setup(m => m.Create(It.IsAny<string>(), It.IsAny<CodeKind>()))
                .Returns((string s, CodeKind ck) => errorListenerMock.Object);
            return errorListenerFactoryMock.Object;
        }

        [Category("Parser")]
        [Test]
        public void TokenStreamParserThrowsPostponedExceptionAfterFinishingFallbackToLlWithoutNonPostponedExceptions()
        {
            var exception = new TokenStreamParserTestException();
            var llErrorListenerFactory = ErrorListenerFactoryWithPostponedException(exception);
            var sllErrorListenerFactory = DummyErrorListenerFactory();

            var llFunction = new Func<IParseTree>(() => null);
            var sllFunction = new Func<IParseTree>(() => throw new Exception());

            var tokenStreamParser = new TestTokenStreamParser(sllErrorListenerFactory, llErrorListenerFactory, llFunction, sllFunction);
            Assert.Throws<TokenStreamParserTestException>(() => tokenStreamParser.Parse("TestModule", null, CodeKind.CodePaneCode, ParserMode.FallBackSllToLl));
        }



        private class TestTokenStreamParser : TokenStreamParserBase
        {
            //Two functions are used because building Func<PredictionMode,IParseTree> instances is cumbersome.
            private readonly Func<IParseTree> _llFunction;
            private readonly Func<IParseTree> _sllFunction;

            public TestTokenStreamParser(
                IParsePassErrorListenerFactory sllErrorListenerFactory,
                IParsePassErrorListenerFactory llErrorListenerFactory,
                Func<IParseTree> llFunction,
                Func<IParseTree> sllFunction)
                : base(sllErrorListenerFactory, llErrorListenerFactory)
            {
                _llFunction = llFunction;
                _sllFunction = sllFunction;
            }

            protected override IParseTree Parse(ITokenStream tokenStream, PredictionMode predictionMode, IParserErrorListener errorListener)
            {
                //Cannot use a switch because PredictionMode.Ll is not a constant.
                if (predictionMode == PredictionMode.Ll)
                {
                    return _llFunction();
                }
                if (predictionMode == PredictionMode.Sll)
                {
                    return _sllFunction();
                }

                throw new System.ComponentModel.InvalidEnumArgumentException();
            }

            protected override void LogAndReset(CommonTokenStream tokenStream, string logWarnMessage, Exception exception)
            { }
        }

        private class TokenStreamParserTestException : Exception
        {}
    }
}