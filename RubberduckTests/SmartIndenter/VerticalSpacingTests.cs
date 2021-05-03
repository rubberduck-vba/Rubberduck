using System.Linq;
using NUnit.Framework;
using Rubberduck.SmartIndenter;
using Rubberduck.UI.Command.ComCommands;
using Rubberduck.VBEditor;
using RubberduckTests.Commands;
using RubberduckTests.Mocks;
using RubberduckTests.Settings;

namespace RubberduckTests.SmartIndenter
{
    [TestFixture]
    public class VerticalSpacingTests
    {
        [Test]
        [Category("Indenter")]
        public void VerticalSpacing_InsertBlankLinesWorks()
        {
            var code = new[]
            {
                "Private Sub TestOne()",
                "End Sub",
                "Private Sub TestTwo()",
                "End Sub",
                "Private Function TestThree()",
                "End Function",
                "Private Function TestFour()",
                "End Function"
            };

            var expected = new[]
            {
                "Private Sub TestOne()",
                "End Sub",
                "",
                "Private Sub TestTwo()",
                "End Sub",
                "",
                "Private Function TestThree()",
                "End Function",
                "",
                "Private Function TestFour()",
                "End Function"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.VerticallySpaceProcedures = true;
                s.LinesBetweenProcedures = 1;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
        public void VerticalSpacing_RemoveExtraLinesWorks()
        {
            var code = new[]
            {
                "Private Sub TestOne()",
                "End Sub",
                "",
                "",
                "",
                "Private Sub TestTwo()",
                "End Sub",
                "",
                "",
                "",
                "Private Function TestThree()",
                "End Function",
                "",
                "",
                "",
                "Private Function TestFour()",
                "End Function"
            };

            var expected = new[]
            {
                "Private Sub TestOne()",
                "End Sub",
                "",
                "Private Sub TestTwo()",
                "End Sub",
                "",
                "Private Function TestThree()",
                "End Function",
                "",
                "Private Function TestFour()",
                "End Function"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.VerticallySpaceProcedures = true;
                s.LinesBetweenProcedures = 1;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
        public void VerticalSpacing_RemoveAllExtraLinesWorks()
        {
            var code = new[]
            {
                "Private Sub TestOne()",
                "End Sub",
                "",
                "",
                "",
                "Private Sub TestTwo()",
                "End Sub",
                "",
                "",
                "",
                "Private Function TestThree()",
                "End Function",
                "",
                "",
                "",
                "Private Function TestFour()",
                "End Function"
            };

            var expected = new[]
            {
                "Private Sub TestOne()",
                "End Sub",
                "Private Sub TestTwo()",
                "End Sub",
                "Private Function TestThree()",
                "End Function",
                "Private Function TestFour()",
                "End Function"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.VerticallySpaceProcedures = true;
                s.LinesBetweenProcedures = 0;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
        public void VerticalSpacing_MixedInsertsAndDeletesWorks()
        {
            var code = new[]
            {
                "Private Sub TestOne()",
                "End Sub",
                "",
                "",
                "",
                "Private Sub TestTwo()",
                "End Sub",
                "Private Function TestThree()",
                "End Function",
                "",
                "",
                "Private Function TestFour()",
                "End Function"
            };

            var expected = new[]
            {
                "Private Sub TestOne()",
                "End Sub",
                "",
                "Private Sub TestTwo()",
                "End Sub",
                "",
                "Private Function TestThree()",
                "End Function",
                "",
                "Private Function TestFour()",
                "End Function"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.VerticallySpaceProcedures = true;
                s.LinesBetweenProcedures = 1;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
        public void VerticalSpacing_InsertBlankLinesWorksWithProperties()
        {
            var code = new[]
            {
                "Function TestFunction() As Long",
                "End Function",
                "Public Property Get TestProperty() As Long",
                "End Property",
                "Public Property Let TestProperty(ByVal foo As Long)",
                "End Property",
                "Sub TestSub()",
                "End Sub"
            };

            var expected = new[]
            {
                "Function TestFunction() As Long",
                "End Function",
                "",
                "Public Property Get TestProperty() As Long",
                "End Property",
                "",
                "Public Property Let TestProperty(ByVal foo As Long)",
                "End Property",
                "",
                "Sub TestSub()",
                "End Sub"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.VerticallySpaceProcedures = true;
                s.LinesBetweenProcedures = 1;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
        public void VerticalSpacing_GroupSingleProperty()
        {
            var code = new[]
            {
                "Function TestFunction() As Long",
                "End Function",
                "Public Property Get TestProperty() As Long",
                "End Property",
                "Public Property Let TestProperty(ByVal foo As Long)",
                "End Property",
                "Sub TestSub()",
                "End Sub"
            };

            var expected = new[]
            {
                "Function TestFunction() As Long",
                "End Function",
                "",
                "Public Property Get TestProperty() As Long",
                "End Property",
                "Public Property Let TestProperty(ByVal foo As Long)",
                "End Property",
                "",
                "Sub TestSub()",
                "End Sub"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.VerticallySpaceProcedures = true;
                s.LinesBetweenProcedures = 1;
                s.GroupRelatedProperties = true;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
        public void VerticalSpacing_GroupMultipleProperties()
        {
            var code = new[]
            {
                "Function TestFunction() As Long",
                "End Function",
                "Public Property Get TestProperty() As Long",
                "End Property",
                "Public Property Let TestProperty(ByVal foo As Long)",
                "End Property",
                "Public Property Get TestProperty2() As Variant",
                "End Property",
                "Public Property Let TestProperty2(ByVal foo As Variant)",
                "End Property",
                "Public Property Set TestProperty2(ByVal foo As Variant)",
                "End Property",
                "Sub TestSub()",
                "End Sub"
            };

            var expected = new[]
            {
                "Function TestFunction() As Long",
                "End Function",
                "",
                "Public Property Get TestProperty() As Long",
                "End Property",
                "Public Property Let TestProperty(ByVal foo As Long)",
                "End Property",
                "",
                "Public Property Get TestProperty2() As Variant",
                "End Property",
                "Public Property Let TestProperty2(ByVal foo As Variant)",
                "End Property",
                "Public Property Set TestProperty2(ByVal foo As Variant)",
                "End Property",
                "",
                "Sub TestSub()",
                "End Sub"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.VerticallySpaceProcedures = true;
                s.LinesBetweenProcedures = 1;
                s.GroupRelatedProperties = true;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
        public void VerticalSpacing_InsertBlankLinesWorksWithTypes()
        {
            var code = new[]
            {
                "Type Foo",
                "End Type",
                "Function TestFunction() As Long",
                "End Function",
                "Sub TestSub()",
                "End Sub"
            };

            var expected = new[]
            {
                "Type Foo",
                "End Type",
                "",
                "Function TestFunction() As Long",
                "End Function",
                "",
                "Sub TestSub()",
                "End Sub"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.VerticallySpaceProcedures = true;
                s.LinesBetweenProcedures = 1;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Ignore("Ignoring this until the indenter is re-written to use the parse tree. Less work than parsing the freak'n code to set up the mocks.")]
        [Category("Indenter")]
        public void VerticalSpacing_IndentProcedureDoesntRemoveSurroundingWhitespace()
        {
            const string inputCode =
                @"Type Foo
    Bar As Long
End Type

Function TestFunction() As Long
TestFunction = 42
End Function

Sub TestSub()
End Sub";

            const string expectedCode =
                @"Type Foo
    Bar As Long
End Type

Function TestFunction() As Long
    TestFunction = 42
End Function

Sub TestSub()
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component, new Selection(6, 1));
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var indentCommand = MockIndenter.ArrangeIndentCurrentProcedureCommand(vbe, new Indenter(vbe.Object, () => IndenterSettingsTests.GetMockIndenterSettings()), state);
                indentCommand.Execute(null);
                Assert.AreEqual(expectedCode, component.CodeModule.Content());
            }
        }

        [Test]
        [Category("Indenter")]
        public void VerticalSpacing_InsertBlankLinesWorksWithEnums()
        {
            var code = new[]
            {
                "Enum Foo",
                "End Enum",
                "Function TestFunction() As Long",
                "End Function",
                "Sub TestSub()",
                "End Sub"
            };

            var expected = new[]
            {
                "Enum Foo",
                "End Enum",
                "",
                "Function TestFunction() As Long",
                "End Function",
                "",
                "Sub TestSub()",
                "End Sub"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.VerticallySpaceProcedures = true;
                s.LinesBetweenProcedures = 1;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
        public void VerticalSpacing_CommentsDoNotEffectSpacing()
        {
            var code = new[]
            {
                "'Comment",
                "Function TestFunction() As Long",
                "End Function",
                "'Comment",
                "Sub TestSub()",
                "End Sub"
            };

            var expected = new[]
            {
                "'Comment",
                "Function TestFunction() As Long",
                "End Function",
                "",
                "'Comment",
                "Sub TestSub()",
                "End Sub"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.VerticallySpaceProcedures = true;
                s.LinesBetweenProcedures = 1;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
        public void VerticalSpacing_DimBlocksStayIntact()
        {
            var code = new[]
            {
                "Private foo As String",
                "Private bar As String",
                "",
                "Function TestFunction() As Long",
                "End Function",
                "Sub TestSub()",
                "End Sub"
            };

            var expected = new[]
            {
                "Private foo As String",
                "Private bar As String",
                "",
                "Function TestFunction() As Long",
                "End Function",
                "",
                "Sub TestSub()",
                "End Sub"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.VerticallySpaceProcedures = true;
                s.LinesBetweenProcedures = 1;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
        public void VerticalSpacing_IgnoresDeclareFunctions()
        {
            var code = new[]
            {
               @"Public Declare Function Foo Lib ""bar.dll"" (X As Any) As Variant",
               @"Public Declare Sub Bar Lib ""bar.dll"" (Y As Integer)",
                "",
                "Function TestFunction() As Long",
                "End Function",
                "Sub TestSub()",
                "End Sub"
            };

            var expected = new[]
            {
               @"Public Declare Function Foo Lib ""bar.dll"" (X As Any) As Variant",
               @"Public Declare Sub Bar Lib ""bar.dll"" (Y As Integer)",
                "",
                "Function TestFunction() As Long",
                "End Function",
                "",
                "Sub TestSub()",
                "End Sub"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.VerticallySpaceProcedures = true;
                s.LinesBetweenProcedures = 1;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
        public void VerticalSpacing_DoesNotChangeSpacingInProcedures()
        {
            var code = new[]
            {
                "Function TestFunction() As Long",
                "",
                "",
                "TestFunction = 42",
                "End Function",
                "Sub TestSub()",
                "Debug.Print TestFunction",
                "",
                "",
                "End Sub"
            };

            var expected = new[]
            {
                "Function TestFunction() As Long",
                "",
                "",
                "    TestFunction = 42",
                "End Function",
                "",
                "Sub TestSub()",
                "    Debug.Print TestFunction",
                "",
                "",
                "End Sub"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.VerticallySpaceProcedures = true;
                s.LinesBetweenProcedures = 1;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
        public void VerticalSpacing_DoesNotChangeSpacingInEnumsOrTypes()
        {
            var code = new[]
            {
                "Enum Foo",
                "",
                "    Bar",
                "    Baz",
                "",
                "",
                "End Enum",
                "",
                "Type Test",
                "",    
                "    MemberOne As String",
                "    MemberTwo As Long",
                "",
                "",    
                "",    
                "End Type",
                "Sub TestSub()",
                "End Sub"
            };

            var expected = new[]
            {
                "Enum Foo",
                "",
                "    Bar",
                "    Baz",
                "",
                "",
                "End Enum",
                "",
                "Type Test",
                "",    
                "    MemberOne As String",
                "    MemberTwo As Long",
                "",
                "",    
                "",    
                "End Type",
                "", 
                "Sub TestSub()",
                "End Sub"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.VerticallySpaceProcedures = true;
                s.LinesBetweenProcedures = 1;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
        public void VerticalSpacing_IgnoredWithIndentProcedureNoSpacing()
        {
            string expected;
            var input = expected =
@"Private Sub TestOne()
End Sub
Private Sub TestTwo()
End Sub
Private Sub TestThree()
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(input, out var component, new Selection(3, 5, 3, 5));
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var indenter = new Indenter(vbe.Object, () =>
                {
                    var s = IndenterSettingsTests.GetMockIndenterSettings();
                    s.VerticallySpaceProcedures = true;
                    s.LinesBetweenProcedures = 1;
                    return s;
                });
                var indentCommand = MockIndenter.ArrangeIndentCurrentProcedureCommand(vbe, indenter, state);
                indentCommand.Execute(null);

                Assert.AreEqual(expected, component.CodeModule.Content());
            }
        }

        [Test]
        [Category("Indenter")]
        public void VerticalSpacing_IgnoredWithIndentProcedureExtraSpacing()
        {
            string expected;
            var input = expected = 
@"Private Sub TestOne()
End Sub


Private Sub TestTwo()
End Sub


Private Sub TestThree()
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(input, out var component, new Selection(3, 5, 3, 5));
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var indenter = new Indenter(vbe.Object, () =>
                {
                    var s = IndenterSettingsTests.GetMockIndenterSettings();
                    s.VerticallySpaceProcedures = true;
                    s.LinesBetweenProcedures = 1;
                    return s;
                });
                var indentCommand = MockIndenter.ArrangeIndentCurrentProcedureCommand(vbe, indenter, state);
                indentCommand.Execute(null);

                Assert.AreEqual(expected, component.CodeModule.Content());
            }
        }

        [Test]
        [Category("Indenter")]
        public void VerticalSpacing_MaintainsSpacingAboveFirstProcedure()
        {
            var code = new[]
            {
                @"Option Explicit",
                "Function TestFunction() As Long",
                "End Function",
                "Sub TestSub()",
                "End Sub"
            };

            var expected = new[]
            {
                @"Option Explicit",
                "",
                "Function TestFunction() As Long",
                "End Function",
                "",
                "Sub TestSub()",
                "End Sub"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.VerticallySpaceProcedures = true;
                s.LinesBetweenProcedures = 1;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
        public void VerticalSpacing_MaintainsIgnoresCommentAboveFirstProcedure()
        {
            var code = new[]
            {
                @"Option Explicit",
                "'Comment",
                "Function TestFunction() As Long",
                "End Function",
                "Sub TestSub()",
                "End Sub"
            };

            var expected = new[]
            {
                @"Option Explicit",
                "",
                "'Comment",
                "Function TestFunction() As Long",
                "End Function",
                "",
                "Sub TestSub()",
                "End Sub"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.VerticallySpaceProcedures = true;
                s.LinesBetweenProcedures = 1;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
        public void VerticalSpacing_RemovesExtraSpacingAboveFirstProcedure()
        {
            var code = new[]
            {
                @"Option Explicit",
                "",
                "",
                "",
                "",
                "Function TestFunction() As Long",
                "End Function",
                "Sub TestSub()",
                "End Sub"
            };

            var expected = new[]
            {
                @"Option Explicit",
                "",
                "Function TestFunction() As Long",
                "End Function",
                "",
                "Sub TestSub()",
                "End Sub"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.VerticallySpaceProcedures = true;
                s.LinesBetweenProcedures = 1;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }
    }
}
