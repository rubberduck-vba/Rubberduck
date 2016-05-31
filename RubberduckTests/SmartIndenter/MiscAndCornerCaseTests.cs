using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.SmartIndenter;

namespace RubberduckTests.SmartIndenter
{
    [TestClass]
    public class MiscAndCornerCaseTests
    {
        [TestMethod]
        [TestCategory("Indenter")]
        public void DeclareFunctionsDoNotIndentNextLine()
        {
            var code = new[]
            {
               @"Public Declare Function Foo Lib ""bar.dll"" (X As Any) As Variant",
               @"Public Declare Sub Bar Lib ""bar.dll"" (Y As Integer)"
            };

            var indenter = new Indenter(null, () => new IndenterSettings());
            var actual = indenter.Indent(code, string.Empty);
            Assert.IsTrue(code.SequenceEqual(actual));
        }

        [TestMethod]
        [TestCategory("Indenter")]
        public void SingleLineIfStatementWorks()
        {
            var code = new[]
            {
                "Public Function Test() As Boolean",
                "If Foo = True Then Bar = False",
                "Test = Bar",
                "End Function"
            };

            var expected = new[]
            {
                "Public Function Test() As Boolean",
                "    If Foo = True Then Bar = False",
                "    Test = Bar",
                "End Function"
            };

            var indenter = new Indenter(null, () => new IndenterSettings());
            var actual = indenter.Indent(code, string.Empty);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [TestMethod]
        [TestCategory("Indenter")]
        public void SingleLineElseIfStatementWorks()
        {
            var code = new[]
            {
                "Public Function Test() As Integer",
                "If Foo = 1 Then",
                "Bar = 3",
                "ElseIf Foo = 2 Then Bar = 2",
                "ElseIf Foo = 3 Then",
                "Bar = 1",
                "End If",
                "Test = Bar",
                "End Function"
            };

            var expected = new[]
            {
                "Public Function Test() As Integer",
                "    If Foo = 1 Then",
                "        Bar = 3",
                "    ElseIf Foo = 2 Then Bar = 2",
                "    ElseIf Foo = 3 Then",
                "        Bar = 1",
                "    End If",
                "    Test = Bar",
                "End Function"
            };

            var indenter = new Indenter(null, () => new IndenterSettings());
            var actual = indenter.Indent(code, string.Empty);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [TestMethod]
        [TestCategory("Indenter")]
        public void LineNumbersAreNotIncludedInIndentAmount()
        {
            var code = new[]
            {
                "Private Sub Test()",
                "5 If Foo Then",
                "10 Debug.Print",
                "15 End If",
                "End Sub"
            };

            var expected = new[]
            {
                "Private Sub Test()",
                "5   If Foo Then",
                "10      Debug.Print",
                "15  End If",
                "End Sub"
            };

            var indenter = new Indenter(null, () => new IndenterSettings());
            var actual = indenter.Indent(code, string.Empty);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [TestMethod]
        [TestCategory("Indenter")]
        public void LineNumberLongerThanIndentFallsBackToOneSpace()
        {
            var code = new[]
            {
                "Private Sub Test()",
                "10000 If Foo Then",
                "10001 Debug.Print",
                "10002 End If",
                "End Sub"
            };

            var expected = new[]
            {
                "Private Sub Test()",
                "10000 If Foo Then",
                "10001   Debug.Print",
                "10002 End If",
                "End Sub"
            };

            var indenter = new Indenter(null, () => new IndenterSettings());
            var actual = indenter.Indent(code, string.Empty);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [TestMethod]
        [TestCategory("Indenter")]
        public void ModulePrecompilerDirectivesWork()
        {
            var code = new[]
            {
                "#Const Foo = Bar",
                "#If Foo Then",
                "Const Baz = 1",
                "#Else",
                "Const Baz = 2",
                "#End If"
            };

            var expected = new[]
            {
                "#Const Foo = Bar",
                "#If Foo Then",
                "    Const Baz = 1",
                "#Else",
                "    Const Baz = 2",
                "#End If"
            };

            var indenter = new Indenter(null, () => new IndenterSettings());
            var actual = indenter.Indent(code, string.Empty);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        //https://github.com/rubberduck-vba/Rubberduck/issues/1286
        [TestMethod]
        [TestCategory("Indenter")]
        public void ReservedWordsInTypesWork()
        {
            var code = new[]
            {
                "Private Type Foo",
                "If As Integer",
                "Select As Integer",
                "For As Integer",
                "Enum As Integer",
                "Type As Integer",
                "Then As Integer",
                "Case As Integer",
                "Function As Integer",
                "Sub As Integer",
                "End Type"
            };

            var expected = new[]
            {
                "Private Type Foo",
                "    If As Integer",
                "    Select As Integer",
                "    For As Integer",
                "    Enum As Integer",
                "    Type As Integer",
                "    Then As Integer",
                "    Case As Integer",
                "    Function As Integer",
                "    Sub As Integer",
                "End Type"
            };

            var indenter = new Indenter(null, () => new IndenterSettings());
            var actual = indenter.Indent(code, string.Empty);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [TestMethod]
        [TestCategory("Indenter")]
        public void IndentationResetsAfterType()
        {
            var code = new[]
            {
                "Private Type Foo",
                "X As Integer",
                "End Type",
                "",
                "Private Sub Test()",
                "If Foo Then",
                "Debug.Print",
                "End If",
                "End Sub"
            };

            var expected = new[]
            {
                "Private Type Foo",
                "    X As Integer",
                "End Type",
                "",
                "Private Sub Test()",
                "    If Foo Then",
                "        Debug.Print",
                "    End If",
                "End Sub"
            };

            var indenter = new Indenter(null, () => new IndenterSettings());
            var actual = indenter.Indent(code, string.Empty);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [TestMethod]
        [TestCategory("Indenter")]
        public void IndentationResetsAfterEnum()
        {
            var code = new[]
            {
                "Public Enum Foo",
                "X = 1",
                "Y = 2",
                "End Enum",
                "",
                "Private Sub Test()",
                "If Foo Then",
                "Debug.Print",
                "End If",
                "End Sub"
            };

            var expected = new[]
            {
                "Public Enum Foo",
                "    X = 1",
                "    Y = 2",
                "End Enum",
                "",
                "Private Sub Test()",
                "    If Foo Then",
                "        Debug.Print",
                "    End If",
                "End Sub"
            };

            var indenter = new Indenter(null, () => new IndenterSettings());
            var actual = indenter.Indent(code, string.Empty);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [TestMethod]
        [TestCategory("Indenter")]
        public void OverIndentationLeftAligns()
        {
            var code = new[]
            {
                "        Private Sub Test()",
                "            If Foo Then",
                "                Debug.Print",
                "            End If",
                "        End Sub"
            };

            var expected = new[]
            {
                "Private Sub Test()",
                "    If Foo Then",
                "        Debug.Print",
                "    End If",
                "End Sub"
            };

            var indenter = new Indenter(null, () => new IndenterSettings());
            var actual = indenter.Indent(code, string.Empty);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }
    }
}
