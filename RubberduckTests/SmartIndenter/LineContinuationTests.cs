using System.Linq;
using NUnit.Framework;
using Rubberduck.SmartIndenter;
using RubberduckTests.Settings;

namespace RubberduckTests.SmartIndenter
{
    [TestFixture]
    public class LineContinuationTests
    {
        [Test]
        [Category("Indenter")]
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
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        // https://github.com/rubberduck-vba/Rubberduck/issues/4795
        [Test]
        [Category("Indenter")]
        public void DeclarationPtrSafeLineAlignsCorrectly()
        {
            var code = new[]
            {
                @"Private Declare PtrSafe Function Foo Lib ""bar.dll"" _",
                 "(x As Long y As Long) As LongPtr"
            };

            var expected = new[]
            {
                @"Private Declare PtrSafe Function Foo Lib ""bar.dll"" _",
                 "                                 (x As Long y As Long) As LongPtr"
            };

            var indenter = new Indenter(null, () => IndenterSettingsTests.GetMockIndenterSettings());
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
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
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
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
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
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
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
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
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
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
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
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
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
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
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        //https://github.com/rubberduck-vba/Rubberduck/issues/1291 
        [Test]
        [Category("Indenter")]
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
            var actual = indenter.Indent(code);
            Assert.IsTrue(code.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
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
            var actual = indenter.Indent(code);
            Assert.IsTrue(code.SequenceEqual(actual));
        }

        [Test]        // Broken in VB6 SmartIndenter.
        [Category("Indenter")]
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

            var expected = new[]
            {
                "Public Sub Test()",
                "    Do _",
                "        While X < 10: X = X + 1: Loop",
                "    Debug.Print X",
                "End Sub"
            };

            var indenter = new Indenter(null, () => IndenterSettingsTests.GetMockIndenterSettings());
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]        // Broken in VB6 SmartIndenter.
        [Category("Indenter")]
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
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
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
                "        + (D * (Bar(E, F, G) / D ^ 2 + Baz(X, Y, Z) / (2 * X)))",
                "End Function"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.IgnoreOperatorsInContinuations = true;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
        public void MultiLineFunctionCallsWorksWithSubFunctionCall()
        {
            var code = new[]
            {
                "Sub Foo()",
                "Dim bar As Long",
                "bar = Baz(Param:=AnotherFunction(Title:=\"foobar\"), _",
                "Behavior:=\"expected\", _",
                "Works:=True)",
                "End Sub"
            };

            var expected = new[]
            {
                "Sub Foo()",
                "    Dim bar As Long",
                "    bar = Baz(Param:=AnotherFunction(Title:=\"foobar\"), _",
                "              Behavior:=\"expected\", _",
                "              Works:=True)",
                "End Sub"
            };

            var indenter = new Indenter(null, () => IndenterSettingsTests.GetMockIndenterSettings());
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
        public void MultiLineFunctionCallsWorksWithSubFunctionCallUnderScoreInParam()
        {
            var code = new[]
            {
                "Sub Foo()",
                "Dim bar As Long",
                "bar = Baz(Param:=AnotherFunction(Title:=FOO_BAR), _",
                "Behavior:=\"expected\", _",
                "Works:=True)",
                "End Sub"
            };

            var expected = new[]
            {
                "Sub Foo()",
                "    Dim bar As Long",
                "    bar = Baz(Param:=AnotherFunction(Title:=FOO_BAR), _",
                "              Behavior:=\"expected\", _",
                "              Works:=True)",
                "End Sub"
            };

            var indenter = new Indenter(null, () => IndenterSettingsTests.GetMockIndenterSettings());
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
        public void MultiLineFunctionCallsWorksWithSubFunctionCallUnderScoreInParamString()
        {
            var code = new[]
            {
                "Sub Foo()",
                "Dim bar As Long",
                "bar = Baz(Param:=AnotherFunction(Title:=\"foo_bar\"), _",
                "Behavior:=\"expected\", _",
                "Works:=True)",
                "End Sub"
            };

            var expected = new[]
            {
                "Sub Foo()",
                "    Dim bar As Long",
                "    bar = Baz(Param:=AnotherFunction(Title:=\"foo_bar\"), _",
                "              Behavior:=\"expected\", _",
                "              Works:=True)",
                "End Sub"
            };

            var indenter = new Indenter(null, () => IndenterSettingsTests.GetMockIndenterSettings());
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
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
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
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
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
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
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        //https://github.com/rubberduck-vba/Rubberduck/issues/1287
        [Test]
        [Category("Indenter")]
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
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        //https://github.com/rubberduck-vba/Rubberduck/issues/1287
        [Test]        // Broken in VB6 SmartIndenter.
        [Category("Indenter")]
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
                "        'Comment _",
                "        as _",
                "        integer",
                "End Sub"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.IndentFirstDeclarationBlock = true;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        //https://github.com/rubberduck-vba/Rubberduck/issues/2407
        [Test] 
        [Category("Indenter")]
        public void ContinuationsInProcedureDeclarationsWithAlignWorks()
        {
            var code = new[]
            {
                "Sub MySub()",
                "Dim x1 As Integer, x2 _",
                "As Integer",
                "End Sub"
            };

            var expected = new[]
            {
                "Sub MySub()",
                "    Dim x1    As Integer, x2 _",
                "              As Integer",
                "End Sub"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.IndentFirstDeclarationBlock = true;
                s.AlignDims = true;
                s.AlignDimColumn = 15;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        //https://github.com/rubberduck-vba/Rubberduck/issues/2407
        [Test] 
        [Category("Indenter")]
        public void ContinuationsInProcedureDeclarationsWithAlignWorksBareType()
        {
            var code = new[]
            {
                "Sub MySub()",
                "Dim x1 As Integer",
                "Dim x2 As _",
                "Integer",
                "Dim x3 As Integer: _",
                "Dim x4 As _",
                "Integer",
                "Dim x5 As Integer _",
                "'Comment _",
                "as _",
                "integer",
                "End Sub"
            };

            var expected = new[]
            {
                "Sub MySub()",
                "    Dim x1    As Integer",
                "    Dim x2    As _",
                "    Integer",
                "    Dim x3    As Integer: _",
                "    Dim x4    As _",
                "    Integer",
                "    Dim x5    As Integer _",
                "        'Comment _",
                "        as _",
                "        integer",
                "End Sub"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.IndentFirstDeclarationBlock = true;
                s.AlignDims = true;
                s.AlignDimColumn = 15;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
        public void ContinuationWithOnlyCommentWorks()
        {
            var code = new[]
            {
                "Sub MySub()",
                "Debug.Print Foo _",
                "'Is this and end of line comment or not...?",
                "End Sub"
            };

            var expected = new[]
            {
                "Sub MySub()",
                "    Debug.Print Foo _",
                "                'Is this and end of line comment or not...?",
                "End Sub"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.AlignCommentsWithCode = false;
                s.IndentFirstDeclarationBlock = true;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        //https://github.com/rubberduck-vba/Rubberduck/issues/1287
        [Test]
        [Category("Indenter")]
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
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        //https://github.com/rubberduck-vba/Rubberduck/issues/1287
        [Test]
        [Category("Indenter")]
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
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        //https://github.com/rubberduck-vba/Rubberduck/issues/1287
        [Test]
        [Category("Indenter")]
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
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        //https://github.com/rubberduck-vba/Rubberduck/issues/1287
        [Test]
        [Category("Indenter")]
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
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        //https://github.com/rubberduck-vba/Rubberduck/issues/2402
        [Test]        // Broken in VB6 SmartIndenter.
        [Category("Indenter")]
        public void SplitNamedParameterAlignsCorrectly()
        {
            var code = new[]
            {
                 "Sub Foo()",
                 "Debug.Print WorksheetFunction.Sum(arg1:=1, arg2:=2, arg3 _",
                 ":=3)",
                 "End Sub"
            };

            var expected = new[]
            {
                 "Sub Foo()",
                 "    Debug.Print WorksheetFunction.Sum(arg1:=1, arg2:=2, arg3 _",
                 "                                      :=3)",
                 "End Sub"
            };

            var indenter = new Indenter(null, () => IndenterSettingsTests.GetMockIndenterSettings());
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        //https://github.com/rubberduck-vba/Rubberduck/issues/2402
        [Test]        // Broken in VB6 SmartIndenter.
        [Category("Indenter")]
        public void MultipleSplitNamedParametersAlignCorrectly()
        {
            var code = new[]
            {
                 "Sub Foo()",
                 "Debug.Print WorksheetFunction.Sum(arg1:=1, arg2:=2, arg3 _",
                 ":=3, arg4:=4, arg5:=6, arg6 _",
                 ":=6)",
                 "End Sub"
            };

            var expected = new[]
            {
                 "Sub Foo()",
                 "    Debug.Print WorksheetFunction.Sum(arg1:=1, arg2:=2, arg3 _",
                 "                                      :=3, arg4:=4, arg5:=6, arg6 _",
                 "                                      :=6)",
                 "End Sub"
            };

            var indenter = new Indenter(null, () => IndenterSettingsTests.GetMockIndenterSettings());
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }
    }
}
