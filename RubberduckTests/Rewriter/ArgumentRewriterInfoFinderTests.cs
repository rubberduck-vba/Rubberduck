using NUnit.Framework;
using RubberduckTests.Mocks;
using System.Linq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter.RewriterInfo;

namespace RubberduckTests.Rewriter
{
    [TestFixture]
    public class ArgumentRewriterInfoFinderTests
    {
        [Test]
        [Category("Rewriter")]
        public void TestArgumentRewriterInfoFinder_Single()
        {
            var inputCode =
                @"Public Sub Foo(fooBar As Long)
End Sub

Public Sub Baz()
    Foo 5
End Sub";
            var enclosingModuleBodyElementName = "Baz";
            var argumentPosition = 1;
            var expectedEnclosedCode =
                @"5";

            TestEnclosedCode(inputCode, enclosingModuleBodyElementName, argumentPosition, expectedEnclosedCode);
        }

        [Test]
        [Category("Rewriter")]
        public void TestArgumentRewriterInfoFinder_EmptyOptional()
        {
            var inputCode =
                @"Public Sub Foo(bar As Long, Optional fooBar As Long = 1, Optional fooBaz As Long = 2)
End Sub

Public Sub Baz()
    Foo 5,, 4
End Sub";
            var enclosingModuleBodyElementName = "Baz";
            var argumentPosition = 2;
            var expectedEnclosedCode =
                @", ";

            TestEnclosedCode(inputCode, enclosingModuleBodyElementName, argumentPosition, expectedEnclosedCode);
        }

        [Test]
        [Category("Rewriter")]
        public void TestArgumentRewriterInfoFindert_First()
        {
            var inputCode =
                @"Public Sub Foo(fooBar As Long, bar As Long, fooBaz As Long)
End Sub

Public Sub Baz()
    Foo 5, 4, 3
End Sub";
            var enclosingModuleBodyElementName = "Baz";
            var argumentPosition = 1;
            var expectedEnclosedCode =
                @"5, ";

            TestEnclosedCode(inputCode, enclosingModuleBodyElementName, argumentPosition, expectedEnclosedCode);
        }

        [Test]
        [Category("Rewriter")]
        public void TestArgumentRewriterInfoFinder_Middle()
        {
            var inputCode =
                @"Public Sub Foo(bar As Long, fooBar As Long, fooBaz As Long)
End Sub

Public Sub Baz()
    Foo 5, 4, 3
End Sub";
            var enclosingModuleBodyElementName = "Baz";
            var argumentPosition = 2;
            var expectedEnclosedCode =
                @"4, ";

            TestEnclosedCode(inputCode, enclosingModuleBodyElementName, argumentPosition, expectedEnclosedCode);
        }

        [Test]
        [Category("Rewriter")]
        public void TestArgumentRewriterInfoFinder_Last()
        {
            var inputCode =
                @"Public Sub Foo(fooBaz As Long, bar As Long, fooBar As Long)
End Sub

Public Sub Baz()
    Foo 5, 4, 3
End Sub";
            var enclosingModuleBodyElementName = "Baz";
            var argumentPosition = 3;
            var expectedEnclosedCode =
                @", 3";

            TestEnclosedCode(inputCode, enclosingModuleBodyElementName, argumentPosition, expectedEnclosedCode);
        }

        private void TestEnclosedCode(string inputCode, string enclosingModuleBodyElementName, int argumentPosition, string expectedEnclosedCode)
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var enclosingModuleBodyElementDeclarationContext = state.DeclarationFinder
                    .MatchName(enclosingModuleBodyElementName).First().Context;
                var enclosingModuleBodyElementContext = enclosingModuleBodyElementDeclarationContext
                    .GetAncestor<VBAParser.ModuleBodyElementContext>();
                var enclosingModuleBodyElementBodyContext =
                    enclosingModuleBodyElementContext.GetDescendent<VBAParser.BlockContext>();
                var argumentListContext =
                    enclosingModuleBodyElementBodyContext.GetDescendent<VBAParser.ArgumentListContext>();
                var argumentContext = argumentListContext.argument(argumentPosition - 1);

                var infoFinder = new ArgumentRewriterInfoFinder();
                var info = infoFinder.GetRewriterInfo(argumentContext);

                var actualEnclosedCode = state.GetCodePaneTokenStream(component.QualifiedModuleName)
                    .GetText(info.StartTokenIndex, info.StopTokenIndex);
                Assert.AreEqual(expectedEnclosedCode, actualEnclosedCode);
            }
        }
    }
}