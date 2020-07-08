using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes;
using Rubberduck.CodeAnalysis.QuickFixes.Concrete;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class RemoveOptionBaseStatementQuickFixTests : QuickFixTestBase
    {
        [Test]
        [Category("QuickFixes")]
        public void OptionBaseZeroStatement_QuickFixWorks_RemoveStatement()
        {
            const string inputCode = "Option Base 0";
            const string expectedCode = "";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new RedundantOptionInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void OptionBaseZeroStatement_QuickFixWorks_RemoveStatement_MultipleLines()
        {
            const string inputCode = @"Option _
Base _
0";

            const string expectedCode = "";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new RedundantOptionInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void OptionBaseZeroStatement_QuickFixWorks_RemoveStatement_InstructionSeparator()
        {
            const string inputCode = "Option Explicit: Option Base 0: Option Base 1";

            const string expectedCode = "Option Explicit: : Option Base 1";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new RedundantOptionInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void OptionBaseZeroStatement_QuickFixWorks_RemoveStatement_InstructionSeparatorAndMultipleLines()
        {
            const string inputCode = @"Option Explicit: Option _
Base _
0: Option Base 1";

            const string expectedCode = "Option Explicit: : Option Base 1";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new RedundantOptionInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }


        protected override IQuickFix QuickFix(RubberduckParserState state)
        {
            return new RemoveOptionBaseStatementQuickFix();
        }
    }
}
