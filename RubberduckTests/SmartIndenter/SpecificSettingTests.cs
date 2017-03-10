using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.SmartIndenter;
using RubberduckTests.Settings;

namespace RubberduckTests.SmartIndenter
{
    [TestClass]
    public class SpecificSettingTests
    {
        [TestMethod]
        [TestCategory("Indenter")]
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

        [TestMethod]
        [TestCategory("Indenter")]
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

        [TestMethod]
        [TestCategory("Indenter")]
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

        [TestMethod]
        [TestCategory("Indenter")]
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

        [TestMethod]
        [TestCategory("Indenter")]
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

        [TestMethod]
        [TestCategory("Indenter")]
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

        [TestMethod]
        [TestCategory("Indenter")]
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

        [TestMethod]
        [TestCategory("Indenter")]
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

        [TestMethod]
        [TestCategory("Indenter")]
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

        [TestMethod]
        [TestCategory("Indenter")]
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

        [TestMethod]
        [TestCategory("Indenter")]
        public void ForceDebugStatementsInColumn1OnWorks()
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
                s.ForceDebugStatementsInColumn1 = true;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [TestMethod]
        [TestCategory("Indenter")]
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

        [TestMethod]
        [TestCategory("Indenter")]
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

        [TestMethod]
        [TestCategory("Indenter")]
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

        [TestMethod]
        [TestCategory("Indenter")]
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

        [TestMethod]
        [TestCategory("Indenter")]
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

        [TestMethod]
        [TestCategory("Indenter")]
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

        [TestMethod]
        [TestCategory("Indenter")]
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

        [TestMethod]
        [TestCategory("Indenter")]
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

        [TestMethod]
        [TestCategory("Indenter")]
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

        [TestMethod]
        [TestCategory("Indenter")]
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

        [TestMethod]
        [TestCategory("Indenter")]
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
        [TestMethod]
        [TestCategory("Indenter")]
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
    }
}
