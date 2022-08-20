using System.Linq;
using NUnit.Framework;
using Rubberduck.SmartIndenter;
using RubberduckTests.Settings;

namespace RubberduckTests.SmartIndenter
{
    [TestFixture]
    public class MultiSegmentLineTests
    {
        [Test]
        [Category("Indenter")]
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

        [Test]
        [Category("Indenter")]
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

        [Test]
        [Category("Indenter")]
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

        [Test]
        [Category("Indenter")]
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

        // https://github.com/rubberduck-vba/Rubberduck/issues/5929
        [Test]        // Broken in VB6 SmartIndenter.
        [Category("Indenter")]
        public void IfThenElseOnSameLineWorks()
        {
            var code = new[]
            {
                "Public Sub Test()",
                "If Foo = 42 Then: Bar = Foo: Else Baz = Bar",
                "Baz = Foo",
                "End Sub"
            };

            var expected = new[]
            {
                "Public Sub Test()",
                "    If Foo = 42 Then: Bar = Foo: Else Baz = Bar",
                "    Baz = Foo",
                "End Sub"
            };

            var indenter = new Indenter(null, () => IndenterSettingsTests.GetMockIndenterSettings());
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
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

        [Test]
        [Category("Indenter")]
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

        // https://github.com/rubberduck-vba/Rubberduck/issues/6007
        [Test]
        [Category("Indenter")]
        public void CommentWithColon()
        {
            var code = new[]
            {
                "Sub foo()",
                "   ' : loop",
                "' bar",
                "End Sub"
            };

            var expected = new[]
            {
                "Sub foo()",
                "    ' : loop",
                "    ' bar",
                "End Sub"
            };

            var indenter = new Indenter(null, () => IndenterSettingsTests.GetMockIndenterSettings()); 
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }
    }
}
