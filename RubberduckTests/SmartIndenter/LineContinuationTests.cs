﻿using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.SmartIndenter;
using RubberduckTests.Settings;

namespace RubberduckTests.SmartIndenter
{
    [TestClass]
    public class LineContinuationTests
    {
        [TestMethod]
        [TestCategory("Indenter")]
        public void DeclarationLineAlignsCorrectly()
        {
            var code = new[]
            {
                @"Public Declare Sub Foo Lib ""bar.dll"" (x As Long, _",
                 "y As Long)"
            };

            var expected = new[]
            {
                @"Public Declare Sub Foo Lib ""bar.dll"" (x As Long, _",
                 "                                      y As Long)"
            };

            var indenter = new Indenter(null, () => IndenterSettingsTests.GetMockIndenterSettings());
            var actual = indenter.Indent(code, string.Empty);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [TestMethod]
        [TestCategory("Indenter")]
        public void FunctionParametersAlignCorrectly()
        {
            var code = new[]
            {
                "Private Sub Test(x As Integer, y As Integer, _",
                "z As Long, abcde As Integer)"
            };

            var expected = new[]
            {
                "Private Sub Test(x As Integer, y As Integer, _",
                "                 z As Long, abcde As Integer)"
            };

            var indenter = new Indenter(null, () => IndenterSettingsTests.GetMockIndenterSettings());
            var actual = indenter.Indent(code, string.Empty);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [TestMethod]
        [TestCategory("Indenter")]
        public void NamedParametersAlignCorrectly()
        {
            var code = new[]
            {
                "Private Sub Foo()",
                "Test X:=1, Y:=2, _",
                "Z:=3, Foobar:=4",
                "End Sub"
            };

            var expected = new[]
            {
                "Private Sub Foo()",
                "    Test X:=1, Y:=2, _",
                "         Z:=3, Foobar:=4",
                "End Sub"
            };

            var indenter = new Indenter(null, () => IndenterSettingsTests.GetMockIndenterSettings());
            var actual = indenter.Indent(code, string.Empty);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [TestMethod]
        [TestCategory("Indenter")]
        public void IgnoreOperatorsInContinuationsOnWorks()
        {
            var code = new[]
            {
                "Private Function Test() As String",
               @"Test = ""String that is "" _",
               @"& ""split into multiple "" _",
               @"& ""lines.""",
                "End Function"
            };

            var expected = new[]
            {
                "Private Function Test() As String",
               @"    Test = ""String that is "" _",
               @"         & ""split into multiple "" _",
               @"         & ""lines.""",
                "End Function"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.IgnoreOperatorsInContinuations = true;
                return s;
            });
            var actual = indenter.Indent(code, string.Empty);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [TestMethod]
        [TestCategory("Indenter")]
        public void IgnoreOperatorsInContinuationsOffWorks()
        {
            var code = new[]
            {
                "Private Function Test() As String",
               @"Test = ""String that is "" _",
               @"& ""split into multiple "" _",
               @"& ""lines.""",
                "End Function"
            };

            var expected = new[]
            {
                "Private Function Test() As String",
               @"    Test = ""String that is "" _",
               @"           & ""split into multiple "" _",
               @"           & ""lines.""",
                "End Function"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.IgnoreOperatorsInContinuations = false;
                return s;
            });
            var actual = indenter.Indent(code, string.Empty);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [TestMethod]
        [TestCategory("Indenter")]
        public void IfThenDoesntGetMangled()
        {
            var code = new[]
            {
                "Private Sub Test()",
                "If (Foo And Bar) Or Baz Then",
                @"Debug.Print ""Foo & Bar | Baz""",
                "End If",
                "End Sub"
            };

            var expected = new[]
            {
                "Private Sub Test()",
                "    If (Foo And Bar) Or Baz Then",
                @"        Debug.Print ""Foo & Bar | Baz""",
                "    End If",
                "End Sub"
            };

            var indenter = new Indenter(null, () => IndenterSettingsTests.GetMockIndenterSettings());
            var actual = indenter.Indent(code, string.Empty);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [TestMethod]
        [TestCategory("Indenter")]
        public void MultilineCommentAlignsCorrectly()
        {
            var code = new[]
            {
                "Private Sub Test()",
                "'Long multiline comment _",
                "blah blah blah, etc. _",
                "etc.",
                "End Sub"
            };

            var expected = new[]
            {
                "Private Sub Test()",
                "    'Long multiline comment _",
                "    blah blah blah, etc. _",
                "    etc.",
                "End Sub"
            };

            var indenter = new Indenter(null, () => IndenterSettingsTests.GetMockIndenterSettings());
            var actual = indenter.Indent(code, string.Empty);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [TestMethod]
        [TestCategory("Indenter")]
        public void ContinuedEndOfLineCommentAligned()
        {
            var code = new[]
            {
                "Private Sub Test()",
                "Dim Foo As Integer 'End of line comment _",
                "Continued end of line comment",
                "End Sub"
            };

            var expected = new[]
            {
                "Private Sub Test()",
                "    Dim Foo As Integer                           'End of line comment _",
                "                                                 Continued end of line comment",
                "End Sub"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.EndOfLineCommentStyle = EndOfLineCommentStyle.AlignInColumn;
                s.EndOfLineCommentColumnSpaceAlignment = 50;
                return s;
            });
            var actual = indenter.Indent(code, string.Empty);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        //https://github.com/rubberduck-vba/Rubberduck/issues/1291 
        [TestMethod]
        [TestCategory("Indenter")]
        public void CommentsWithLineContinuationsWork()
        {
            var code = new[]
            {
                "'Private Sub Test()",
                "'Long multiline comment _",
                "'blah blah blah, etc. _",
                "'etc.",
                "'End Sub"
            };

            var indenter = new Indenter(null, () => IndenterSettingsTests.GetMockIndenterSettings());
            var actual = indenter.Indent(code, string.Empty);
            Assert.IsTrue(code.SequenceEqual(actual));
        }

        [TestMethod]
        [TestCategory("Indenter")]
        public void RemsWithLineContinuationsWork()
        {
            var code = new[]
            {
                "'Private Sub Test()",
                "Rem Long multiline comment _",
                "Rem blah blah blah, etc. _",
                "Rem etc.",
                "'End Sub"
            };

            var indenter = new Indenter(null, () => IndenterSettingsTests.GetMockIndenterSettings());
            var actual = indenter.Indent(code, string.Empty);
            Assert.IsTrue(code.SequenceEqual(actual));
        }

        [TestMethod, Ignore]    // Broken in VB6 SmartIndenter. Should be same fix as IfThenElseOnSameLineWorks()
        [TestCategory("Indenter")]
        public void DoWhileOnTwoLinesWorks()
        {
            var code = new[]
            {
                "Public Sub Test()",
                "Do _",
                "While X < 10: X = X + 1: Loop",
                "Debug.Print X",
                "End Sub"
            };

            //TODO: Not sure if this is what should be expected...
            var expected = new[]
            {
                "Public Sub Test()",
                "    Do _",
                "    While X < 10: X = X + 1: Loop",
                "    Debug.Print X",
                "End Sub"
            };

            var indenter = new Indenter(null, () => IndenterSettingsTests.GetMockIndenterSettings());
            var actual = indenter.Indent(code, string.Empty);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [TestMethod] // Broken in VB6 SmartIndenter.
        [TestCategory("Indenter")]
        public void ContinuedIfThenWorks()
        {
            var code = new[]
            {
                "Public Sub Test()",
                "If _",
                "Foo = True _",
                "Then",
                "Debug.Print",
                "End If",
                "End Sub"
            };

            var expected = new[]
            {
                "Public Sub Test()",
                "    If _",
                "        Foo = True _",
                "              Then",
                "        Debug.Print",
                "    End If",
                "End Sub"
            };

            var indenter = new Indenter(null, () => IndenterSettingsTests.GetMockIndenterSettings());
            var actual = indenter.Indent(code, string.Empty);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [TestMethod]
        [TestCategory("Indenter")]
        public void MultipleFunctionCallsWorks()
        {
            var code = new[]
            {
                "Private Function Foo()",
                "Foo = Bar(A, B, C + 1) = D * Baz(E, F, G) / (4 * X * Y) * Bar(A + 1, B + 1, C) _",
                "+ (D * (Bar(E, F, G) / D ^ 2 + Baz(X, Y, Z) / (2 * X)))",
                "End Function"
            };

            var expected = new[]
            {
                "Private Function Foo()",
                "    Foo = Bar(A, B, C + 1) = D * Baz(E, F, G) / (4 * X * Y) * Bar(A + 1, B + 1, C) _",
                "          + (D * (Bar(E, F, G) / D ^ 2 + Baz(X, Y, Z) / (2 * X)))",
                "End Function"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.IgnoreOperatorsInContinuations = false;
                return s;
            });
            var actual = indenter.Indent(code, string.Empty);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [TestMethod]
        [TestCategory("Indenter")]
        public void UncontinuedNestedFunctionWorks()
        {
            var code = new[]
            {
                "Public Sub Test()",
                "Debug.Print Foo(X, Bar(Y, Z), _",
                "Z)",
                "End Sub"
            };

            var expected = new[]
            {
                "Public Sub Test()",
                "    Debug.Print Foo(X, Bar(Y, Z), _",
                "                    Z)",
                "End Sub"
            };

            var indenter = new Indenter(null, () => IndenterSettingsTests.GetMockIndenterSettings());
            var actual = indenter.Indent(code, string.Empty);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [TestMethod]
        [TestCategory("Indenter")]
        public void ContinuedNestedFunctionWorks()
        {
            var code = new[]
            {
                "Public Sub Test()",
                "Debug.Print Foo(X, Bar(A, B, _",
                "C), Z)",
                "End Sub"
            };

            var expected = new[]
            {
                "Public Sub Test()",
                "    Debug.Print Foo(X, Bar(A, B, _",
                "                           C), Z)",
                "End Sub"
            };

            var indenter = new Indenter(null, () => IndenterSettingsTests.GetMockIndenterSettings());
            var actual = indenter.Indent(code, string.Empty);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [TestMethod]
        [TestCategory("Indenter")]
        public void MultiLineContinuedFunctionWorks()
        {
            var code = new[]
            {
                "Public Sub Test()",
                "Debug.Print Foo(Var1, Var2, _",
                "Var3, Var4, _",
                "Var5)",
                "End Sub"
            };

            var expected = new[]
            {
                "Public Sub Test()",
                "    Debug.Print Foo(Var1, Var2, _",
                "                    Var3, Var4, _",
                "                    Var5)",
                "End Sub"
            };

            var indenter = new Indenter(null, () => IndenterSettingsTests.GetMockIndenterSettings());
            var actual = indenter.Indent(code, string.Empty);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        //https://github.com/rubberduck-vba/Rubberduck/issues/1287
        [TestMethod]
        [TestCategory("Indenter")]
        public void ContinuationsInTypeDeclarationsWorks()
        {
            var code = new[]
            {
                "Private Type MyType",
                "x1 As Integer",
                "x2 _",
                "As Integer",
                "x3 As _",
                "Integer",
                "x4 _",
                "As _",
                "Integer",
                "x5 As Integer: _",
                "x6 As _",
                "Integer",
                "x7 As Integer _",
                "'Comment _",
                "as _",
                "integer",
                "End Type"
            };

            var expected = new[]
            {
                "Private Type MyType",
                "    x1 As Integer",
                "    x2 _",
                "    As Integer",
                "    x3 As _",
                "    Integer",
                "    x4 _",
                "    As _",
                "    Integer",
                "    x5 As Integer: _",
                "    x6 As _",
                "    Integer",
                "    x7 As Integer _",
                "    'Comment _",
                "    as _",
                "    integer",
                "End Type"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.AlignCommentsWithCode = true;
                return s;
            });
            var actual = indenter.Indent(code, string.Empty);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        //https://github.com/rubberduck-vba/Rubberduck/issues/1287
        [TestMethod, Ignore]    // Broken in VB6 SmartIndenter.
        [TestCategory("Indenter")]
        public void ContinuationsInProcedureDeclarationsWorks()
        {
            var code = new[]
            {
                "Sub MySub()",
                "Dim x1 As Integer",
                "Dim _",
                "x2 _",
                "As Integer",
                "Dim x3 As _",
                "Integer",
                "Dim x4 _",
                "As _",
                "Integer",
                "Dim x5 As Integer: _",
                "Dim x6 As _",
                "Integer",
                "Dim x7 As Integer _",
                "'Comment _",
                "as _",
                "integer",
                "End Sub"
            };

            //TODO: Not sure if this is what should be expected...
            var expected = new[]
            {
                "Sub MySub()",
                "    Dim x1 As Integer",
                "    Dim _",
                "    x2 _",
                "    As Integer",
                "    Dim x3 As _",
                "    Integer",
                "    Dim x4 _",
                "    As _",
                "    Integer",
                "    Dim x5 As Integer: _",
                "    Dim x6 As _",
                "    Integer",
                "    Dim x7 As Integer _",
                "    'Comment _",
                "    as _",
                "    integer",
                "End Sub"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.AlignCommentsWithCode = true;
                s.IndentFirstDeclarationBlock = true;
                return s;
            });
            var actual = indenter.Indent(code, string.Empty);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        //https://github.com/rubberduck-vba/Rubberduck/issues/1287
        [TestMethod, Ignore]    // Broken in VB6 SmartIndenter.
        [TestCategory("Indenter")]
        public void ContinuationsInProcedureDeclarationsNoCommentAlignWorks()
        {
            var code = new[]
            {
                "Sub MySub()",
                "Dim x1 As Integer",
                "Dim _",
                "x2 _",
                "As Integer",
                "Dim x3 As _",
                "Integer",
                "Dim x4 _",
                "As _",
                "Integer",
                "Dim x5 As Integer: _",
                "Dim x6 As _",
                "Integer",
                "Dim x7 As Integer _",
                "'Comment _",
                "as _",
                "integer",
                "End Sub"
            };

            //TODO: Not sure if this is what should be expected...
            var expected = new[]
            {
                "Sub MySub()",
                "    Dim x1 As Integer",
                "    Dim _",
                "    x2 _",
                "    As Integer",
                "    Dim x3 As _",
                "    Integer",
                "    Dim x4 _",
                "    As _",
                "    Integer",
                "    Dim x5 As Integer: _",
                "    Dim x6 As _",
                "    Integer",
                "    Dim x7 As Integer _",
                "'Comment _",
                "as _",
                "integer",
                "End Sub"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.AlignCommentsWithCode = false;
                s.IndentFirstDeclarationBlock = true;
                return s;
            });
            var actual = indenter.Indent(code, string.Empty);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        //https://github.com/rubberduck-vba/Rubberduck/issues/1287
        [TestMethod]
        [TestCategory("Indenter")]
        public void TrailingMultiLineCommentFirstWorks()
        {
            var code = new[]
            {
                "Sub MySub1()",
                "Dim x7 As Integer _",
                "'Comment foo _",
                "",
                "End Sub",
                "",
                "Sub MySub2()",
                "Dim x7 As Integer _",
                "'Comment foo _",
                "foo",
                "End Sub"
            };

            var expected = new[]
            {
                "Sub MySub1()",
                "    Dim x7 As Integer _",
                "        'Comment foo _",
                "",
                "End Sub",
                "",
                "Sub MySub2()",
                "    Dim x7 As Integer _",
                "        'Comment foo _",
                "        foo",
                "End Sub"
            };

            var indenter = new Indenter(null, () => IndenterSettingsTests.GetMockIndenterSettings());
            var actual = indenter.Indent(code, string.Empty);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        //https://github.com/rubberduck-vba/Rubberduck/issues/1287
        [TestMethod]
        [TestCategory("Indenter")]
        public void TrailingMultiLineCommentSecondWorks()
        {
            var code = new[]
            {
                "Sub MySub2()",
                "Dim x7 As Integer _",
                "'Comment foo _",
                "foo",
                "End Sub",
                "",
                "Sub MySub1()",
                "Dim x7 As Integer _",
                "'Comment foo _",
                "",
                "End Sub"
            };

            var expected = new[]
            {
                "Sub MySub2()",
                "    Dim x7 As Integer _",
                "        'Comment foo _",
                "        foo",
                "End Sub",
                "",
                "Sub MySub1()",
                "    Dim x7 As Integer _",
                "        'Comment foo _",
                "",
                "End Sub"
            };

            var indenter = new Indenter(null, () => IndenterSettingsTests.GetMockIndenterSettings());
            var actual = indenter.Indent(code, string.Empty);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        //https://github.com/rubberduck-vba/Rubberduck/issues/1287
        [TestMethod]
        [TestCategory("Indenter")]
        public void TrailingMultiLineCommentFirstNoAlignWorks()
        {
            var code = new[]
            {
                "Sub MySub1()",
                "Dim x7 As Integer _",
                "'Comment foo _",
                "",
                "End Sub",
                "",
                "Sub MySub2()",
                "Dim x7 As Integer _",
                "'Comment foo _",
                "foo",
                "End Sub"
            };

            var expected = new[]
            {
                "Sub MySub1()",
                "    Dim x7 As Integer _",
                "    'Comment foo _",
                "",
                "End Sub",
                "",
                "Sub MySub2()",
                "    Dim x7 As Integer _",
                "    'Comment foo _",
                "    foo",
                "End Sub"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.AlignContinuations = false;
                return s;
            });
            var actual = indenter.Indent(code, string.Empty);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        //https://github.com/rubberduck-vba/Rubberduck/issues/1287
        [TestMethod]
        [TestCategory("Indenter")]
        public void TrailingMultiLineCommentSecondNoAlignWorks()
        {
            var code = new[]
            {
                "Sub MySub2()",
                "Dim x7 As Integer _",
                "'Comment foo _",
                "foo",
                "End Sub",
                "",
                "Sub MySub1()",
                "Dim x7 As Integer _",
                "'Comment foo _",
                "",
                "End Sub"
            };

            var expected = new[]
            {
                "Sub MySub2()",
                "    Dim x7 As Integer _",
                "    'Comment foo _",
                "    foo",
                "End Sub",
                "",
                "Sub MySub1()",
                "    Dim x7 As Integer _",
                "    'Comment foo _",
                "",
                "End Sub"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.AlignContinuations = false;
                return s;
            });
            var actual = indenter.Indent(code, string.Empty);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }
    }
}
