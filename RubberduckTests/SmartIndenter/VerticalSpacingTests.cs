using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.SmartIndenter;
using RubberduckTests.Settings;

namespace RubberduckTests.SmartIndenter
{
    [TestClass]
    public class VerticalSpacingTests
    {
        [TestMethod]
        [TestCategory("Indenter")]
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

        [TestMethod]
        [TestCategory("Indenter")]
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

        [TestMethod]
        [TestCategory("Indenter")]
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

        [TestMethod]
        [TestCategory("Indenter")]
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

        [TestMethod]
        [TestCategory("Indenter")]
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

        [TestMethod]
        [TestCategory("Indenter")]
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

        [TestMethod]
        [TestCategory("Indenter")]
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

        [TestMethod]
        [TestCategory("Indenter")]
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

        [TestMethod]
        [TestCategory("Indenter")]
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

        [TestMethod]
        [TestCategory("Indenter")]
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

        [TestMethod]
        [TestCategory("Indenter")]
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

        [TestMethod]
        [TestCategory("Indenter")]
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
    }
}
