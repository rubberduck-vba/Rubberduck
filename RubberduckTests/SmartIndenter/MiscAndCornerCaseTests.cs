using System.Linq;
using NUnit.Framework;
using Rubberduck.SmartIndenter;
using Rubberduck.UI.Command;
using RubberduckTests.Settings;

namespace RubberduckTests.SmartIndenter
{
    [TestFixture]
    public class MiscAndCornerCaseTests
    {
        [Test]
        [Category("Indenter")]
        public void LowerCaseKeywordNext()
        {
            var code = new[]
            {
                "Dim i As Long",
                "Dim maxRow As Long",
                "",
                "maxRow = 120",
                "For i = 20 To maxRow Step 19",
                "",
                "Sheet5.Range(Rows(i), Rows(i + 4)).Cut Sheet61.Range(Rows(i))",
                "next i"
            };

            var output = new[]
            {
                "Dim i As Long",
                "Dim maxRow As Long",
                "",
                "maxRow = 120",
                "For i = 20 To maxRow Step 19",
                "",
                "    Sheet5.Range(Rows(i), Rows(i + 4)).Cut Sheet61.Range(Rows(i))",
                "next i"
            };

            var indenter = new Indenter(null, () => IndenterSettingsTests.GetMockIndenterSettings());
            var actual = indenter.Indent(code);
            Assert.IsTrue(output.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
        public void DeclareFunctionsDoNotIndentNextLine()
        {
            var code = new[]
            {
                @"Public Declare Function Foo Lib ""bar.dll"" (X As Any) As Variant",
                @"Public Declare Sub Bar Lib ""bar.dll"" (Y As Integer)"
            };

            var indenter = new Indenter(null, () => IndenterSettingsTests.GetMockIndenterSettings());
            var actual = indenter.Indent(code);
            Assert.IsTrue(code.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
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

            var indenter = new Indenter(null, () => IndenterSettingsTests.GetMockIndenterSettings());
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
        public void ElseIfStatementWorks()
        {
            var code = new[]
            {
                "Public Function Test() As Integer",
                "If Foo = 1 Then",
                "Bar = 3",
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
                "    ElseIf Foo = 3 Then",
                "        Bar = 1",
                "    End If",
                "    Test = Bar",
                "End Function"
            };

            var indenter = new Indenter(null, () => IndenterSettingsTests.GetMockIndenterSettings());
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        //https://github.com/rubberduck-vba/Rubberduck/issues/1858
        [Test]
        [Category("Indenter")]
        public void MultipleElseIfStatementWorks()
        {
            var code = new[]
            {
                "Public Sub Test()",
                "If Foo And Bar Then",
                "Call Foobar",
                "ElseIf Not Foo Then",
                "Call Baz",
                "ElseIf Not Bar Then",
                "Call NoBaz",
                "Else",
                "MsgBox \"No Foos or Bars\"",
                "End If",
                "End Sub"
            };

            var expected = new[]
            {
                "Public Sub Test()",
                "    If Foo And Bar Then",
                "        Call Foobar",
                "    ElseIf Not Foo Then",
                "        Call Baz",
                "    ElseIf Not Bar Then",
                "        Call NoBaz",
                "    Else",
                "        MsgBox \"No Foos or Bars\"",
                "    End If",
                "End Sub"
            };

            var indenter = new Indenter(null, () => IndenterSettingsTests.GetMockIndenterSettings());
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        //https://github.com/rubberduck-vba/Rubberduck/issues/1858
        //https://github.com/rubberduck-vba/Rubberduck/issues/2233
        [Test]
        [Category("Indenter")]
        public void IfThenBareElseStatementWorks()
        {
            var code = new[]
            {
                "Public Sub Test()",
                "If Foo And Bar Then Foobar Else",
                "Baz",
                "End Sub"
            };

            var expected = new[]
            {
                "Public Sub Test()",
                "    If Foo And Bar Then Foobar Else",
                "    Baz",
                "End Sub"
            };

            var indenter = new Indenter(null, () => IndenterSettingsTests.GetMockIndenterSettings());
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        //https://github.com/rubberduck-vba/Rubberduck/issues/1858
        [Test]
        [Category("Indenter")]
        public void SingleLineElseIfElseStatementWorks()
        {
            var code = new[]
            {
                "Public Sub Test()",
                "If x Then",
                "NotFoobar",
                "ElseIf Foo And Bar Then Foobar",
                "Else",
                "Baz",
                "End If",
                "End Sub"
            };

            var expected = new[]
            {
                "Public Sub Test()",
                "    If x Then",
                "        NotFoobar",
                "    ElseIf Foo And Bar Then Foobar",
                "    Else",
                "        Baz",
                "    End If",
                "End Sub"
            };

            var indenter = new Indenter(null, () => IndenterSettingsTests.GetMockIndenterSettings());
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
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

            var indenter = new Indenter(null, () => IndenterSettingsTests.GetMockIndenterSettings());
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
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

            var indenter = new Indenter(null, () => IndenterSettingsTests.GetMockIndenterSettings());
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
        public void NegativeLineNumbersWork()
        {
            var code = new[]
            {
                "Private Sub Test()",
                "-5 If Foo Then",
                "-10 Debug.Print",
                "-15 End If",
                "End Sub"
            };

            var expected = new[]
            {
                "Private Sub Test()",
                "-5  If Foo Then",
                "-10     Debug.Print",
                "-15 End If",
                "End Sub"
            };

            var indenter = new Indenter(null, () => IndenterSettingsTests.GetMockIndenterSettings());
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
        public void HexLineNumbersWork()
        {
            var code = new[]
            {
                "Private Sub Test()",
                "&HAAA If Foo Then",
                "&HABE Debug.Print",
                "&HAD2 End If",
                "End Sub"
            };

            var expected = new[]
            {
                "Private Sub Test()",
                "2730 If Foo Then",
                "2750    Debug.Print",
                "2770 End If",
                "End Sub"
            };

            var indenter = new Indenter(null, () => IndenterSettingsTests.GetMockIndenterSettings());
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
        public void LineNumbersWithColonWork()
        {
            var code = new[]
           {
                "Private Sub Test()",
                "5: If Foo Then",
                "10: Debug.Print",
                "15: End If",
                "End Sub"
            };

            var expected = new[]
            {
                "Private Sub Test()",
                "5:  If Foo Then",
                "10:     Debug.Print",
                "15: End If",
                "End Sub"
            };

            var indenter = new Indenter(null, () => IndenterSettingsTests.GetMockIndenterSettings());
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
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

            var indenter = new Indenter(null, () => IndenterSettingsTests.GetMockIndenterSettings());
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
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

            var indenter = new Indenter(null, () => IndenterSettingsTests.GetMockIndenterSettings());
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        //https://github.com/rubberduck-vba/Rubberduck/issues/1286
        [Test]
        [Category("Indenter")]
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

            var indenter = new Indenter(null, () => IndenterSettingsTests.GetMockIndenterSettings());
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
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

            var indenter = new Indenter(null, () => IndenterSettingsTests.GetMockIndenterSettings());
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
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

            var indenter = new Indenter(null, () => IndenterSettingsTests.GetMockIndenterSettings());
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
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

            var indenter = new Indenter(null, () => IndenterSettingsTests.GetMockIndenterSettings());
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        //http://chat.stackexchange.com/transcript/message/33575758#33575758
        [Test]                // Broken in VB6 SmartIndenter.
        [Category("Indenter")]
        public void SubFooTokenIsNotInterpretedAsProcedureStart()
        {
            var code = new[]
            {
                "Public Sub Test()",
                "If Subject = 0 Then",
                "Subject = 1",
                "End If",
                "End Sub"
            };

            var expected = new[]
            {
                "Public Sub Test()",
                "    If Subject = 0 Then",
                "        Subject = 1",
                "    End If",
                "End Sub"
            };

            var indenter = new Indenter(null, () => IndenterSettingsTests.GetMockIndenterSettings());
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        //https://github.com/rubberduck-vba/Rubberduck/issues/2133
        [Test]                // Broken in VB6 SmartIndenter.
        [Category("Indenter")]  
        public void MultiLineDimWithCommasDontAlignDimWorks()
        {
            var code = new[]
            {
                "Public Sub FooBar()",
                "Dim foo As Boolean, bar As String _",
                ", baz As String _",
                ", somethingElse As String",
                "Dim x As Integer",
                "If Not foo Then",
                "x = 1",
                "End If",
                "End Sub"
            };

            var expected = new[]
            {
                "Public Sub FooBar()",
                "    Dim foo As Boolean, bar As String _",
                "    , baz As String _",
                "    , somethingElse As String",
                "    Dim x As Integer",
                "    If Not foo Then",
                "        x = 1",
                "    End If",
                "End Sub"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.AlignDims = false;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        //https://github.com/rubberduck-vba/Rubberduck/issues/2133
        [Test]                // Broken in VB6 SmartIndenter.
        [Category("Indenter")]
        public void MultiLineDimWithCommasAlignDimsWorks()
        {
            var code = new[]
            {
                "Public Sub FooBar()",
                "Dim foo As Boolean, bar As String _",
                ", baz As String _",
                ", somethingElse As String",
                "Dim x As Integer",
                "If Not foo Then",
                "x = 1",
                "End If",
                "End Sub"
            };

            var expected = new[]
            {
                "Public Sub FooBar()",
                "    Dim foo   As Boolean, bar As String _",
                "    , baz     As String _",
                "    , somethingElse As String",
                "    Dim x     As Integer",
                "    If Not foo Then",
                "        x = 1",
                "    End If",
                "End Sub"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.AlignDims = true;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        //https://github.com/rubberduck-vba/Rubberduck/issues/2133
        [Test]                // Broken in VB6 SmartIndenter.
        [Category("Indenter")]
        public void MultiLineDimWithCommasDontIndentFirstBlockWorks()
        {
            var code = new[]
            {
                "Public Sub FooBar()",
                "Dim foo As Boolean, bar As String _",
                ", baz As String _",
                ", somethingElse As String",
                "Dim x As Integer",
                "If Not foo Then",
                "x = 1",
                "End If",
                "End Sub"
            };

            var expected = new[]
            {
                "Public Sub FooBar()",
                "Dim foo As Boolean, bar As String _",
                ", baz As String _",
                ", somethingElse As String",
                "Dim x As Integer",
                "    If Not foo Then",
                "        x = 1",
                "    End If",
                "End Sub"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.IndentFirstDeclarationBlock = false;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]       
        [Category("Indenter")]
        public void QuotesInsideStringLiteralsWork()
        {
            var code = new[]
            {
                "Public Sub Test()",
                "Debug.Print \"This is a \"\" in the middle of a string.\"",
                "End Sub"
            };

            var expected = new[]
            {
                "Public Sub Test()",
                "    Debug.Print \"This is a \"\" in the middle of a string.\"",
                "End Sub"
            };

            var indenter = new Indenter(null, () => IndenterSettingsTests.GetMockIndenterSettings());
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
        public void SingleQuoteInEndOfLineCommentWorks()
        {
            var code = new[]
            {
                "Public Sub Test()",
                "Debug.Print Chr$(34) 'Prints a single '\"' character.",
                "End Sub"
            };

            var expected = new[]
            {
                "Public Sub Test()",
                "    Debug.Print Chr$(34)                         'Prints a single '\"' character.",
                "End Sub"
            };

            var indenter = new Indenter(null, () => IndenterSettingsTests.GetMockIndenterSettings());
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        //http://chat.stackexchange.com/transcript/message/33933348#33933348
        [Test]        // Broken in VB6 SmartIndenter. Also broken in the code's conception. Sheesh - keep the cat off the keyboard.
        [Category("Indenter")]
        public void BracketedIdentifiersWork()
        {
            var code = new[]
            {
                "Sub test()",
                "Dim _",
                "s _",
                "( _",
                "[Option _ Explicit] _",
                "+ _",
                "1 _",
                "To _",
                "( _",
                "[Evil : \"\" Comment \" 'here] _",
                ") _",
                "+ _",
                "[End _ Sub] _",
                ") _",
                "As _",
                "String _",
                "* _",
                "25",
                "End Sub"
            };

            var expected = new[]
            {
                "Sub test()",
                "    Dim _",
                "    s _",
                "    ( _",
                "    [Option _ Explicit] _",
                "    + _",
                "    1 _",
                "    To _",
                "    ( _",
                "    [Evil : \"\" Comment \" 'here] _",
                "    ) _",
                "    + _",
                "    [End _ Sub] _",
                "    ) _",
                "    As _",
                "    String _",
                "    * _",
                "    25",
                "End Sub"
            };

            var indenter = new Indenter(null, () => IndenterSettingsTests.GetMockIndenterSettings());
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        //https://github.com/rubberduck-vba/Rubberduck/issues/2696
        [Test]
        // Broken in VB6 SmartIndenter.
        [Category("Indenter")]
        public void BracketsInEndOfLineCommentsWork()
        {
            var code = new[]
            {
                "Public Sub Test()",
                "Debug.Print \"foo\" \'update [foo].[bar] in the frob.",
                "End Sub"
            };

            var expected = new[]
            {
                "Public Sub Test()",
                "    Debug.Print \"foo\"                            'update [foo].[bar] in the frob.",
                "End Sub"
            };

            var indenter = new Indenter(null, () => IndenterSettingsTests.GetMockIndenterSettings());
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        //https://github.com/rubberduck-vba/Rubberduck/issues/2604
        [Test]
        [Category("Indenter")]
        public void AlignmentAnchorsInStringLiteralsAreIgnored()
        {
            var code = new[]
            {
                "Sub Test()",
                "Dim LoremIpsum As String",
               @"LoremIpsum = ""Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nam dictum,"" & vbCrLf _",
               @"& ""felis in tempor finibus, arcu lectus molestie urna, eget interdum turpis"" & vbCrLf _",
               @"& ""tellus ac diam. Nulla mauris lectus, vulputate et fringilla ac, iaculis eget urna."" & vbCrLf _",
               @"& ""Ut feugiat felis lacinia eros vestibulum facilisis. Ut euismod dapibus augue,"" & vbCrLf _",
               @"& ""lacinia elementum elit dictum in. Nam in imperdiet tortor. Curabitur efficitur libero"" & vbCrLf _",
               @"& ""lacus, et placerat metus sodales sit amet.""",
                "Debug.Print LoremIpsum",
                "End Sub"
            };

            var expected = new[]
            {
                "Sub Test()",
                "    Dim LoremIpsum As String",
               @"    LoremIpsum = ""Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nam dictum,"" & vbCrLf _",
               @"               & ""felis in tempor finibus, arcu lectus molestie urna, eget interdum turpis"" & vbCrLf _",
               @"               & ""tellus ac diam. Nulla mauris lectus, vulputate et fringilla ac, iaculis eget urna."" & vbCrLf _",
               @"               & ""Ut feugiat felis lacinia eros vestibulum facilisis. Ut euismod dapibus augue,"" & vbCrLf _",
               @"               & ""lacinia elementum elit dictum in. Nam in imperdiet tortor. Curabitur efficitur libero"" & vbCrLf _",
               @"               & ""lacus, et placerat metus sodales sit amet.""",
                "    Debug.Print LoremIpsum",
                "End Sub"
            };

            var indenter = new Indenter(null, () => IndenterSettingsTests.GetMockIndenterSettings());
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
        public void ElseWithEndingColonWorks()
        {
            var code = new[]
            {
                "Sub Foo()",
                "If True Then",
                "Debug.Print \"True\"",
                "Else:",
                "Debug.Print \"False\"",
                "End If",
                "End Sub"
            };

            var expected = new[]
            {
                "Sub Foo()",
                "    If True Then",
                "        Debug.Print \"True\"",
                "    Else:",
                "        Debug.Print \"False\"",
                "    End If",
                "End Sub"
            };

            var indenter = new Indenter(null, () => IndenterSettingsTests.GetMockIndenterSettings());
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }
        
        [Test]
        [Category("Indenter")]
        public void ElseWithTrailingSegmentWorks()
        {
            var code = new[]
            {
                "Sub Foo()",
                "If True Then",
                "Debug.Print \"True\"",
                "Else: Debug.Print \"False\"",
                "End If",
                "End Sub"
            };

            var expected = new[]
            {
                "Sub Foo()",
                "    If True Then",
                "        Debug.Print \"True\"",
                "    Else: Debug.Print \"False\"",
                "    End If",
                "End Sub"
            };

            var indenter = new Indenter(null, () => IndenterSettingsTests.GetMockIndenterSettings());
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        //failing test for https://github.com/rubberduck-vba/Rubberduck/issues/3210
        [Test]
        [Ignore("Most likely requires the parse tree.")]
        [Category("Indenter")]
        public void LineNumbersInsideContinuationsWork()
        {
            var code = new[]
            {
                "Sub Foo()",
                " _",
                "10",
                " _",
                "foo _",
                ": Beep",
                "",
                "20 bar: Beep",
                "End Sub"
            };

            var expected = new[]
            {
                "Sub Foo()",
                "   _",
                "10",
                "   _",
                "foo _",
                ":  Beep",
                "",
                "20 bar: Beep",
                "End Sub"
            };

            var indenter = new Indenter(null, () => IndenterSettingsTests.GetMockIndenterSettings());
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
        public void ReplacementPatternsInStringLiteralWorks()
        {
            var code = new[]
            {
                "Public Sub Test()",
                "Debug.Print \"a*${test}b\"",
                "End Sub"
            };

            var expected = new[]
            {
                "Public Sub Test()",
                "    Debug.Print \"a*${test}b\"",
                "End Sub"
            };

            var indenter = new Indenter(null, () => IndenterSettingsTests.GetMockIndenterSettings());
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }
    }
}
