using System.Linq;
using NUnit.Framework;
using Rubberduck.SmartIndenter;
using RubberduckTests.Settings;

namespace RubberduckTests.SmartIndenter
{
    [TestFixture]
    public class SpecificSettingTests
    {
        [Test]
        [Category("Indenter")]
        public void IndentEntireProcedureBodyOffWorks()
        {
            var code = new[]
            {
                "Private Sub Test()",
                "If Foo Then",
                "Debug.Print",
                "End If",
                "End Sub"
            };

            var expected = new[]
            {
                "Private Sub Test()",
                "If Foo Then",
                "    Debug.Print",
                "End If",
                "End Sub"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.IndentEntireProcedureBody = false;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
        public void IndentEntireProcedureBodyOnWorks()
        {
            var code = new[]
            {
                "Private Sub Test()",
                "If Foo Then",
                "Debug.Print",
                "End If",
                "End Sub"
            };

            var expected = new[]
            {
                "Private Sub Test()",
                "    If Foo Then",
                "        Debug.Print",
                "    End If",
                "End Sub"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.IndentEntireProcedureBody = true;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
        public void IndentFirstCommentBlockOffWorks()
        {
            var code = new[]
            {
                "Public Function Test() As String",
                "'Comment block",
                "'Comment block",
                "Dim Foo As Long",
                "Dim Bar As Long",
                @"Test = ""Passed!""",
                "End Function"
            };

            var expected = new[]
            {
                "Public Function Test() As String",
                "'Comment block",
                "'Comment block",
                "    Dim Foo As Long",
                "    Dim Bar As Long",
               @"    Test = ""Passed!""",
                "End Function"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.IndentFirstCommentBlock = false;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
        public void IndentFirstCommentBlockOnWorks()
        {
            var code = new[]
            {
                "Public Function Test() As String",
                "'Comment block",
                "'Comment block",
                "Dim Foo As Long",
                "Dim Bar As Long",
                @"Test = ""Passed!""",
                "End Function"
            };

            var expected = new[]
            {
                "Public Function Test() As String",
                "    'Comment block",
                "    'Comment block",
                "    Dim Foo As Long",
                "    Dim Bar As Long",
               @"    Test = ""Passed!""",
                "End Function"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.IndentFirstCommentBlock = true;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
        public void IndentEnumTypeCommentBlockOffWorks()
        {
            var code = new[]
            {
                "Public Enum Test",
                "'Comment block",
                "'Comment block",
                "Foo",
                "Bar",
                "End Enum"
            };

            var expected = new[]
            {
                "Public Enum Test",
                "    'Comment block",
                "    'Comment block",
                "    Foo",
                "    Bar",
                "End Enum"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.IndentEnumTypeAsProcedure = false;
                s.IndentFirstCommentBlock = true;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
        public void IndentEnumTypeCommentBlockOnWorks()
        {
            var code = new[]
            {
                "Public Enum Test",
                "'Comment block",
                "'Comment block",
                "Foo",
                "Bar",
                "End Enum"
            };

            var expected = new[]
            {
                "Public Enum Test",
                "'Comment block",
                "'Comment block",
                "    Foo",
                "    Bar",
                "End Enum"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.IndentEnumTypeAsProcedure = true;
                s.IndentFirstCommentBlock = false;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
        public void IndentFirstCommentBlockOffOnlyOnFirst()
        {
            var code = new[]
            {
                "Public Function Test() As String",
                "'Comment block",
                "'Comment block",
                "Dim Foo As Long",
                "Dim Bar As Long",
                @"Test = ""Passed!""",
                "'Not in comment block",
                "End Function"
            };

            var expected = new[]
            {
                "Public Function Test() As String",
                "'Comment block",
                "'Comment block",
                "    Dim Foo As Long",
                "    Dim Bar As Long",
               @"    Test = ""Passed!""",
                "    'Not in comment block",
                "End Function"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.IndentFirstCommentBlock = false;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
        public void IndentFirstCommentBlockOffIgnoreEmptyOnlyOnFirst()
        {
            var code = new[]
            {
                "Public Function Test() As String",
                "'Comment block",
                "",
                "'Comment block",
                "Dim Foo As Long",
                "Dim Bar As Long",
                @"Test = ""Passed!""",
                "'Not in comment block",
                "End Function"
            };

            var expected = new[]
            {
                "Public Function Test() As String",
                "'Comment block",
                "",
                "'Comment block",
                "    Dim Foo As Long",
                "    Dim Bar As Long",
                @"    Test = ""Passed!""",
                "    'Not in comment block",
                "End Function"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.IndentFirstCommentBlock = false;
                s.IgnoreEmptyLinesInFirstBlocks = true;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
        public void IndentFirstDeclarationBlockOffWorks()
        {
            var code = new[]
            {
                "Public Function Test() As String",
                "'Comment block",
                "'Comment block",
                "Dim Foo As Long",
                "Dim Bar As Long",
                @"Test = ""Passed!""",
                "End Function"
            };

            var expected = new[]
            {
                "Public Function Test() As String",
                "    'Comment block",
                "    'Comment block",
                "Dim Foo As Long",
                "Dim Bar As Long",
               @"    Test = ""Passed!""",
                "End Function"
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
        public void IndentFirstDeclarationBlockOffIgnoreEmptyWorks()
        {
            var code = new[]
            {
                "Public Function Test() As String",
                "'Comment block",
                "'Comment block",
                "Dim Foo As Long",
                "",
                "Dim Bar As Long",
                @"Test = ""Passed!""",
                "End Function"
            };

            var expected = new[]
            {
                "Public Function Test() As String",
                "    'Comment block",
                "    'Comment block",
                "Dim Foo As Long",
                "",
                "Dim Bar As Long",
                @"    Test = ""Passed!""",
                "End Function"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.IndentFirstDeclarationBlock = false;
                s.IgnoreEmptyLinesInFirstBlocks = true;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
        public void IndentFirstDeclarationCommentBlockOffWorksIntermingled()
        {
            var code = new[]
            {
                "Public Function Test() As String",
                "'Foo comment",
                "Dim Foo As Long",
                "'Bar comment",
                "Dim Bar As Long",
                @"Test = ""Passed!""",
                "End Function"
            };

            var expected = new[]
            {
                "Public Function Test() As String",
                "'Foo comment",
                "Dim Foo As Long",
                "'Bar comment",
                "Dim Bar As Long",
                @"    Test = ""Passed!""",
                "End Function"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.IndentFirstDeclarationBlock = false;
                s.IndentFirstCommentBlock = false;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
        public void IndentFirstDeclarationCommentBlockOffIgnoreEmptyWorksIntermingled()
        {
            var code = new[]
            {
                "Public Function Test() As String",
                "'Foo comment",
                "Dim Foo As Long",
                "",
                "'Bar comment",
                "Dim Bar As Long",
                @"Test = ""Passed!""",
                "End Function"
            };

            var expected = new[]
            {
                "Public Function Test() As String",
                "'Foo comment",
                "Dim Foo As Long",
                "",
                "'Bar comment",
                "Dim Bar As Long",
                @"    Test = ""Passed!""",
                "End Function"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.IndentFirstDeclarationBlock = false;
                s.IndentFirstCommentBlock = false;
                s.IgnoreEmptyLinesInFirstBlocks = true;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
        public void IndentFirstDeclarationBlockOnWorks()
        {
            var code = new[]
            {
                "Public Function Test() As String",
                "'Comment block",
                "'Comment block",
                "Dim Foo As Long",
                "Dim Bar As Long",
                @"Test = ""Passed!""",
                "End Function"
            };

            var expected = new[]
            {
                "Public Function Test() As String",
                "    'Comment block",
                "    'Comment block",
                "    Dim Foo As Long",
                "    Dim Bar As Long",
               @"    Test = ""Passed!""",
                "End Function"
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

        [Test]
        [Category("Indenter")]
        public void IndentFirstDeclarationBlockOffOnlyOnFirst()
        {
            var code = new[]
            {
                "Public Function Test() As String",
                "'Comment block",
                "'Comment block",
                "Dim Foo As Long",
                "Dim Bar As Long",
                @"Test = ""Passed!""",
                "Dim Baz as Long",
                "End Function"
            };

            var expected = new[]
            {
                "Public Function Test() As String",
                "    'Comment block",
                "    'Comment block",
                "Dim Foo As Long",
                "Dim Bar As Long",
               @"    Test = ""Passed!""",
                "    Dim Baz as Long",
                "End Function"
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
        public void ForceDebugPrintInColumn1OnWorks()
        {
            var code = new[]
            {
                "Private Sub Test()",
                "If Foo Then",
                "Debug.Print",
                "End If",
                "End Sub"
            };

            var expected = new[]
            {
                "Private Sub Test()",
                "    If Foo Then",
                "Debug.Print",
                "    End If",
                "End Sub"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.ForceDebugPrintInColumn1 = true;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
        public void ForceDebugAssertInColumn1OnWorks()
        {
            var code = new[]
            {
                "Private Sub Test()",
                "If Foo Then",
                "Debug.Assert False",
                "End If",
                "End Sub"
            };

            var expected = new[]
            {
                "Private Sub Test()",
                "    If Foo Then",
                "Debug.Assert False",
                "    End If",
                "End Sub"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.ForceDebugAssertInColumn1 = true;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
        public void ForceStopInColumn1OnWorks()
        {
            var code = new[]
            {
                "Private Sub Test()",
                "If Foo Then",
                "Stop",
                "End If",
                "End Sub"
            };

            var expected = new[]
            {
                "Private Sub Test()",
                "    If Foo Then",
                "Stop",
                "    End If",
                "End Sub"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.ForceStopInColumn1 = true;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
        public void ForceOnlyDebugPrintInColumn1OnWorks()
        {
            var code = new[]
            {
                "Private Sub Test()",
                "If Foo Then",
                "Debug.Print \"Foo\"",
                "Debug.Assert False",
                "Stop",
                "End If",
                "End Sub"
            };

            var expected = new[]
            {
                "Private Sub Test()",
                "    If Foo Then",
                "Debug.Print \"Foo\"",
                "        Debug.Assert False",
                "        Stop",
                "    End If",
                "End Sub"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.ForceDebugPrintInColumn1 = true;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
        public void ForceOnlyDebugAssertInColumn1OnWorks()
        {
            var code = new[]
            {
                "Private Sub Test()",
                "If Foo Then",
                "Debug.Print \"Foo\"",
                "Debug.Assert False",
                "Stop",
                "End If",
                "End Sub"
            };

            var expected = new[]
            {
                "Private Sub Test()",
                "    If Foo Then",
                "        Debug.Print \"Foo\"",
                "Debug.Assert False",
                "        Stop",
                "    End If",
                "End Sub"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.ForceDebugAssertInColumn1 = true;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
        public void ForceOnlyStopInColumn1OnWorks()
        {
            var code = new[]
            {
                "Private Sub Test()",
                "If Foo Then",
                "Debug.Print \"Foo\"",
                "Debug.Assert False",
                "Stop",
                "End If",
                "End Sub"
            };

            var expected = new[]
            {
                "Private Sub Test()",
                "    If Foo Then",
                "        Debug.Print \"Foo\"",
                "        Debug.Assert False",
                "Stop",
                "    End If",
                "End Sub"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.ForceStopInColumn1 = true;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
        public void EnablingForceDebugStatementsEnablesAllSubOptions()
        {
            var options = new IndenterSettings
            {
                ForceDebugStatementsInColumn1 = true
            };
            
            Assert.IsTrue(options.ForceDebugPrintInColumn1);
            Assert.IsTrue(options.ForceDebugAssertInColumn1);
            Assert.IsTrue(options.ForceStopInColumn1);
        }

        [Test]
        [Category("Indenter")]
        public void DisablingForceDebugStatementsDisablesAllSubOptions()
        {
            var options = new IndenterSettings
            {
                ForceDebugPrintInColumn1 = true,
                ForceDebugAssertInColumn1 = true,
                ForceStopInColumn1 = true,               
            };

            options.ForceDebugStatementsInColumn1 = false;

            Assert.IsFalse(options.ForceDebugPrintInColumn1);
            Assert.IsFalse(options.ForceDebugAssertInColumn1);
            Assert.IsFalse(options.ForceStopInColumn1);
        }

        [Test]
        [Category("Indenter")]
        public void AlignCommentsWithCodeOnWorks()
        {
            var code = new[]
            {
                "Private Sub Test()",
                "If Foo Then",
                "  'Comment",
                "Debug.Print",
                "End If",
                "End Sub"
            };

            var expected = new[]
            {
                "Private Sub Test()",
                "    If Foo Then",
                "        'Comment",
                "        Debug.Print",
                "    End If",
                "End Sub"
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

        [Test]
        [Category("Indenter")]
        public void AlignCommentsWithCodeOffWorks()
        {
            var code = new[]
            {
                "Private Sub Test()",
                "If Foo Then",
                "  'Comment",
                "Debug.Print",
                "End If",
                "End Sub"
            };

            var expected = new[]
            {
                "Private Sub Test()",
                "    If Foo Then",
                "  'Comment",
                "        Debug.Print",
                "    End If",
                "End Sub"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.AlignCommentsWithCode = false;
                s.IndentFirstCommentBlock = false;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
        public void ForceDebugStatementsInColumn1OffWorks()
        {
            var code = new[]
            {
                "Private Sub Test()",
                "If Foo Then",
                "Debug.Print",
                "End If",
                "End Sub"
            };

            var expected = new[]
            {
                "Private Sub Test()",
                "    If Foo Then",
                "        Debug.Print",
                "    End If",
                "End Sub"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.ForceDebugStatementsInColumn1 = false;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
        public void ForceCompilerDirectivesInColumn1OffWorks()
        {
            var code = new[]
            {
                "Private Sub Test()",
                "#If Foo Then",
                "Debug.Print",
                "#End If",
                "End Sub"
            };

            var expected = new[]
            {
                "Private Sub Test()",
                "    #If Foo Then",
                "        Debug.Print",
                "    #End If",
                "End Sub"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.ForceCompilerDirectivesInColumn1 = false;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
        public void ForceCompilerDirectivesInColumn1OnWorks()
        {
            var code = new[]
            {
                "Private Sub Test()",
                "#If Foo Then",
                "Debug.Print",
                "#End If",
                "End Sub"
            };

            var expected = new[]
            {
                "Private Sub Test()",
                "#If Foo Then",
                "        Debug.Print",
                "#End If",
                "End Sub"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.ForceCompilerDirectivesInColumn1 = true;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
        public void IndentCaseOffWorks()
        {
            var code = new[]
            {
                "Public Function Test() As Integer",
                "Select Case Foo",
                "Case Bar",
                "Test = 1",
                "Case Baz",
                "Test = 2",
                "End Select",
                "End Function"
            };

            var expected = new[]
            {
                "Public Function Test() As Integer",
                "    Select Case Foo",
                "    Case Bar",
                "        Test = 1",
                "    Case Baz",
                "        Test = 2",
                "    End Select",
                "End Function"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.IndentCase = false;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
        public void IndentCaseOnWorks()
        {
            var code = new[]
            {
                "Public Function Test() As Integer",
                "Select Case Foo",
                "Case Bar",
                "Test = 1",
                "Case Baz",
                "Test = 2",
                "End Select",
                "End Function"
            };

            var expected = new[]
            {
                "Public Function Test() As Integer",
                "    Select Case Foo",
                "        Case Bar",
                "            Test = 1",
                "        Case Baz",
                "            Test = 2",
                "    End Select",
                "End Function"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.IndentCase = true;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
        public void IndentCompilerDirectivesOnWorks()
        {
            var code = new[]
            {
                "Private Sub Test()",
                "#If Foo Then",
                "Debug.Print",
                "#End If",
                "End Sub"
            };

            var expected = new[]
            {
                "Private Sub Test()",
                "    #If Foo Then",
                "        Debug.Print",
                "    #End If",
                "End Sub"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.IndentCompilerDirectives = true;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
        public void IndentCompilerDirectivesOffWorks()
        {
            var code = new[]
            {
                "Private Sub Test()",
                "#If Foo Then",
                "Debug.Print",
                "#End If",
                "End Sub"
            };

            var expected = new[]
            {
                "Private Sub Test()",
                "    #If Foo Then",
                "    Debug.Print",
                "    #End If",
                "End Sub"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.IndentCompilerDirectives = false;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
        public void AlignDimsOnWorks()
        {
            var code = new[]
            {
                "Private Sub Test()",
                "Dim Foo As Integer",
                "Dim Bar As Long",
                "End Sub"
            };

            var expected = new[]
            {
                "Private Sub Test()",
                "    Dim Foo   As Integer",
                "    Dim Bar   As Long",
                "End Sub"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.AlignDims = true;
                s.AlignDimColumn = 15;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
        public void AlignDimsFallsBackToOneSpace()
        {
            var code = new[]
            {
                "Private Sub Test()",
                "Dim Foo As Integer",
                "Dim LongVariableName As Long",
                "End Sub"
            };

            var expected = new[]
            {
                "Private Sub Test()",
                "    Dim Foo   As Integer",
                "    Dim LongVariableName As Long",
                "End Sub"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.AlignDims = true;
                s.AlignDimColumn = 15;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        //https://github.com/rubberduck-vba/Rubberduck/issues/1290
        [Test]
        [Category("Indenter")]
        public void IndentSpacesSettingIsUsed()
        {
            var code = new[]
            {
                "Private Sub Test()",
                "If Foo Then",
                "Debug.Print",
                "End If",
                "End Sub"
            };

            var expected = new[]
            {
                "Private Sub Test()",
                "   If Foo Then",
                "      Debug.Print",
                "   End If",
                "End Sub"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.IndentSpaces = 3;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        //https://github.com/rubberduck-vba/Rubberduck/issues/3800
        [Test]
        [Category("Indenter")]
        public void AlignDimColumnWorksVariedSpacing()
        {
            var code = new[]
            {
                "Private Sub Test()",
                "' comment",
                "    Dim a                                       As String",
                "    Dim b As Object",
                "    Dim c                               As Long",
                "End Sub"
            };

            var expected = new[]
            {
                "Private Sub Test()",
                "' comment",
                "    Dim a                                       As String",
                "    Dim b                                       As Object",
                "    Dim c                                       As Long",
                "End Sub"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.IndentSpaces = 4;
                s.IndentFirstCommentBlock = false;
                s.AlignDims = true;
                s.AlignDimColumn = 49;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        //https://github.com/rubberduck-vba/Rubberduck/issues/3800
        [Test]
        [Category("Indenter")]
        public void AlignDimColumnWorksAlreadyAligned()
        {
            string[] expected;
            var code = expected = new[]
            {
                "Private Sub Test()",
                "' comment",
                "    Dim a                                       As String",
                "    Dim b                                       As Object",
                "    Dim c                                       As Long",
                "End Sub"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.IndentSpaces = 4;
                s.IndentFirstCommentBlock = false;
                s.AlignDims = true;
                s.AlignDimColumn = 49;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
        public void IgnoreEmptyLinesWorks()
        {
            var code = new[]
            {
                "Private Sub Test()",
                "If Foo Then",
                "Debug.Print",
                "     ",
                "End If",
                "End Sub"
            };

            var expected = new[]
            {
                "Private Sub Test()",
                "    If Foo Then",
                "        Debug.Print",
                "     ",
                "    End If",
                "End Sub"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.EmptyLineHandlingMethod = EmptyLineHandling.Ignore;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }


        [Test]
        [Category("Indenter")]
        public void RemoveEmptyLinesWorks()
        {
            var code = new[]
            {
                "Private Sub Test()",
                "If Foo Then",
                "Debug.Print",
                "     ",
                "End If",
                "End Sub"
            };

            var expected = new[]
            {
                "Private Sub Test()",
                "    If Foo Then",
                "        Debug.Print",
                "",
                "    End If",
                "End Sub"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.EmptyLineHandlingMethod = EmptyLineHandling.Remove;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
        public void IndentEmptyLinesWorks()
        {
            var code = new[]
            {
                "Private Sub Test()",
                "If Foo Then",
                "Debug.Print",
                "     ",
                "End If",
                "End Sub"
            };

            var expected = new[]
            {
                "Private Sub Test()",
                "    If Foo Then",
                "        Debug.Print",
                "        ",
                "    End If",
                "End Sub"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.EmptyLineHandlingMethod = EmptyLineHandling.Indent;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }
    }
}
