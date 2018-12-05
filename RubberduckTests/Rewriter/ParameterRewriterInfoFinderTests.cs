using NUnit.Framework;
using RubberduckTests.Mocks;
using System.Linq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Rewriter.RewriterInfo;

namespace RubberduckTests.Rewriter
{
    [TestFixture]
    public class ParameterRewriterInfoFinderTests
    {
        [Test]
        [Category("Rewriter")]
        public void TestParameterRewriterInfoFinder_Single()
        {
            var inputCode =
                @"Public Sub Foo(fooBar As Long)
End Sub";
            var parameterName = "fooBar";
            var expectedEnclosedCode =
                @"fooBar As Long";

            TestEnclosedCode(inputCode, parameterName, expectedEnclosedCode);
        }

        [Test]
        [Category("Rewriter")]
        public void TestParameterRewriterInfoFinder_SingleOptional()
        {
            var inputCode =
                @"Public Sub Foo(Optional fooBar As Long = 1)
End Sub";
            var parameterName = "fooBar";
            var expectedEnclosedCode =
                @"Optional fooBar As Long = 1";

            TestEnclosedCode(inputCode, parameterName, expectedEnclosedCode);
        }

        [Test]
        [Category("Rewriter")]
        public void TestParameterRewriterInfoFindert_First()
        {
            var inputCode =
                @"Public Sub Foo(fooBar As Long, bar As Long, baz As String)
End Sub";
            var parameterName = "fooBar";
            var expectedEnclosedCode =
                @"fooBar As Long, ";

            TestEnclosedCode(inputCode, parameterName, expectedEnclosedCode);
        }

        [Test]
        [Category("Rewriter")]
        public void TestParameterRewriterInfoFinder_Middle()
        {
            var inputCode =
                @"Public Sub Foo(bar As Long, fooBar As Long, baz As String)
End Sub";
            var parameterName = "fooBar";
            var expectedEnclosedCode =
                @"fooBar As Long, ";

            TestEnclosedCode(inputCode, parameterName, expectedEnclosedCode);
        }

        [Test]
        [Category("Rewriter")]
        public void TestParameterRewriterInfoFinder_Last()
        {
            var inputCode =
                @"Public Sub Foo(baz As Long, bar As Long, fooBar As String)
End Sub";
            var parameterName = "fooBar";
            var expectedEnclosedCode =
                @", fooBar As String";

            TestEnclosedCode(inputCode, parameterName, expectedEnclosedCode);
        }

        private void TestEnclosedCode(string inputCode, string parameterName, string expectedEnclosedCode)
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var argContext = state.DeclarationFinder.MatchName(parameterName).First().Context;

                var infoFinder = new ParameterRewriterInfoFinder();
                var info = infoFinder.GetRewriterInfo(argContext);

                var actualEnclosedCode = state.GetCodePaneTokenStream(component.QualifiedModuleName)
                    .GetText(info.StartTokenIndex, info.StopTokenIndex);
                Assert.AreEqual(expectedEnclosedCode, actualEnclosedCode);
            }
        }
    }
}
