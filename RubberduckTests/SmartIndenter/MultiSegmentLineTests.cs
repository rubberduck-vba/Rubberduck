using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.SmartIndenter;
using RubberduckTests.Settings;

namespace RubberduckTests.SmartIndenter
{
    [TestClass]
    public class MultiSegmentLineTests
    {
        [TestMethod]
        [TestCategory("Indenter")]
        public void SingleLineFunctionsNotIndented()
        {
            var code = new[]
            {
                "Private Function Foo(): Foo = 42: End Function",
                "Private Sub Bar(): Debug.Assert Foo = 42: End Sub"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.VerticallySpaceProcedures = false;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(code.SequenceEqual(actual));
        }

        [TestMethod]
        [TestCategory("Indenter")]
        public void SingleLineEnumsNotIndented()
        {
            var code = new[]
            {
                "Public Enum Foo: X = 1: Y = 2: End Enum",
                "Public Enum Bar: X = 1: Y = 2: End Enum"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.VerticallySpaceProcedures = false;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(code.SequenceEqual(actual));
        }

        [TestMethod]
        [TestCategory("Indenter")]
        public void SingleLineTypesNotIndented()
        {
            var code = new[]
            {
                "Public Type Foo: X As Integer: End Type",
                "Public Enum Bar: X = 1: Y = 2: End Enum"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.VerticallySpaceProcedures = false;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(code.SequenceEqual(actual));
        }

        [TestMethod]
        [TestCategory("Indenter")]
        public void UnmatchedEnumsNotIndent()
        {
            var code = new[]
            {
                "Public Enum Foo: X = 1: Y = 2: End Enum: Public Enum Bar",
                "X = 1: Y = 2: End Enum"
            };

            var indenter = new Indenter(null, () => IndenterSettingsTests.GetMockIndenterSettings());
            var actual = indenter.Indent(code);
            Assert.IsTrue(code.SequenceEqual(actual));
        }

        [TestMethod]        // Broken in VB6 SmartIndenter.
        [TestCategory("Indenter")]
        public void IfThenElseOnSameLineWorks()
        {
            var code = new[]
            {
                "Public Sub Test()",
                "If Foo = 42 Then: Bar = Foo: Else",
                "Baz = Foo",
                "End If",
                "End Sub"
            };

            var expected = new[]
            {
                "Public Sub Test()",
                "    If Foo = 42 Then: Bar = Foo: Else",
                "        Baz = Foo",
                "    End If",
                "End Sub"
            };

            var indenter = new Indenter(null, () => IndenterSettingsTests.GetMockIndenterSettings());
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [TestMethod]
        [TestCategory("Indenter")]
        public void MixedSelectSyntaxWorks()
        {
            var code = new[]
            {
                "Public Sub Test()",
                "Select Case Foo",
               @"Case 1: Debug.Print ""Foo""",
                "Case 2",
               @"Debug.Print ""Bar""",
                "End Select",
                "End Sub"
            };

            var expected = new[]
            {
                "Public Sub Test()",
                "    Select Case Foo",
               @"        Case 1: Debug.Print ""Foo""",
                "        Case 2",
               @"            Debug.Print ""Bar""",
                "    End Select",
                "End Sub"
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
        public void UnfinishedCaseOnNextLineWorks()
        {
            var code = new[]
            {
                "Public Sub Test()",
                "Select Case Foo",
               @"Case 1: Debug.Print ""Foo"": Case 2",
               @"Debug.Print ""Bar""",
                "End Select",
                "End Sub"
            };

            var expected = new[]
            {
                "Public Sub Test()",
                "    Select Case Foo",
               @"        Case 1: Debug.Print ""Foo"": Case 2",
               @"            Debug.Print ""Bar""",
                "    End Select",
                "End Sub"
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
    }
}
