using System.Linq;
using NUnit.Framework;
using Rubberduck.SmartIndenter;
using RubberduckTests.Settings;

namespace RubberduckTests.SmartIndenter
{
    [TestFixture]
    public class EndOfLineCommentTests
    {
        [Test]
        [Category("Indenter")]
        public void AbsolutePositionWorks()
        {
            var code = new[]
            {
                "Private Sub Test()",
                "Debug.Print        'Comment",
                "End Sub"
            };

            var expected = new[]
            {
                "Private Sub Test()",
                "    Debug.Print    'Comment",
                "End Sub"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.EndOfLineCommentStyle = EndOfLineCommentStyle.Absolute;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
        public void SameGapWorks()
        {
            var code = new[]
            {
                "Private Sub Test()",
                "Debug.Print   'Comment",
                "End Sub"
            };

            var expected = new[]
            {
                "Private Sub Test()",
                "    Debug.Print   'Comment",
                "End Sub"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.EndOfLineCommentStyle = EndOfLineCommentStyle.SameGap;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
        public void StandardGapWorks()
        {
            var code = new[]
            {
                "Private Sub Test()",
                "Debug.Print 'Comment",
                "End Sub"
            };

            var expected = new[]
            {
                "Private Sub Test()",
                "    Debug.Print        'Comment",
                "End Sub"
            };


            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.EndOfLineCommentStyle = EndOfLineCommentStyle.StandardGap;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
        public void AlignInColumnWorks()
        {
            var code = new[]
            {
                "Private Sub Test()",
                "Debug.Print 'Comment",
                "End Sub"
            };

            var expected = new[]
            {
                "Private Sub Test()",
                "    Debug.Print                                  'Comment",
                "End Sub"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.EndOfLineCommentStyle = EndOfLineCommentStyle.AlignInColumn;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
        public void AlignInColumnFallsBackToOneSpace()
        {
            var code = new[]
            {
                "Private Sub Test()",
               @"Debug.Print ""This string ends with a comment."" 'Comment",
                "End Sub"
            };

            var expected = new[]
            {
                "Private Sub Test()",
               @"    Debug.Print ""This string ends with a comment."" 'Comment",
                "End Sub"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.EndOfLineCommentStyle = EndOfLineCommentStyle.AlignInColumn;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
        public void AbsoluteFallsBackToOneSpace()
        {
            var code = new[]
            {
                "Private Sub Test()",
                "Debug.Print  'Comment",
                "End Sub"
            };

            var expected = new[]
            {
                "Private Sub Test()",
                "    Debug.Print 'Comment",
                "End Sub"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.EndOfLineCommentStyle = EndOfLineCommentStyle.Absolute;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
        public void CommentOnlyLineIgnoresEndOfLineSetting()
        {
            var code = new[]
            {
                "Private Sub Test()",
                "'Comment",
                "End Sub"
            };

            var expected = new[]
            {
                "Private Sub Test()",
                "    'Comment",
                "End Sub"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.EndOfLineCommentStyle = EndOfLineCommentStyle.AlignInColumn;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
        public void WorksOutsideOfProcedures()
        {
            var code = new[]
            {
                "#Const Foo = Bar 'Comment",
                "Private Sub Test()",
                "End Sub"
            };

            var expected = new[]
            {
                "#Const Foo = Bar                                 'Comment",
                "",
                "Private Sub Test()",                
                "End Sub"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.EndOfLineCommentStyle = EndOfLineCommentStyle.AlignInColumn;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Indenter")]
        public void WorksInsideEnums()
        {
            var code = new[]
            {
                "Public Enum Foo",
                "X = 1 'Comment",
                "Y = 2",
                "End Enum"
            };

            var expected = new[]
            {
                "Public Enum Foo",
                "    X = 1                                        'Comment",
                "    Y = 2",
                "End Enum"
            };

            var indenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.EndOfLineCommentStyle = EndOfLineCommentStyle.AlignInColumn;
                return s;
            });
            var actual = indenter.Indent(code);
            Assert.IsTrue(expected.SequenceEqual(actual));
        }
    }
}
