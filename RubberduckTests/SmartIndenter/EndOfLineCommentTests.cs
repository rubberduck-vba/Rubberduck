using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.SmartIndenter;
using RubberduckTests.Settings;

namespace RubberduckTests.SmartIndenter
{
    [TestClass]
    public class EndOfLineCommentTests
    {
        [TestMethod]
        [TestCategory("Indenter")]
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

        [TestMethod]
        [TestCategory("Indenter")]
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

        [TestMethod]
        [TestCategory("Indenter")]
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

        [TestMethod]
        [TestCategory("Indenter")]
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

        [TestMethod]
        [TestCategory("Indenter")]
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

        [TestMethod]
        [TestCategory("Indenter")]
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

        [TestMethod]
        [TestCategory("Indenter")]
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

        [TestMethod]
        [TestCategory("Indenter")]
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

        [TestMethod]
        [TestCategory("Indenter")]
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
