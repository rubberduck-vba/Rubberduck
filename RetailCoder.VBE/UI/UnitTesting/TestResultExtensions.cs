using System.Drawing;
using System.Runtime.InteropServices;
using Rubberduck.UnitTesting;

namespace Rubberduck.UI.UnitTesting
{
    public static class TestResultExtensions
    {
        public static Image Icon(this TestResult result)
        {
            var image = Properties.Resources.question_white;
            if (result != null)
            {
                switch (result.Outcome)
                {
                    case TestOutcome.Succeeded:
                        image = Properties.Resources.tick_circle;
                        break;

                    case TestOutcome.Failed:
                        image = Properties.Resources.cross_circle;
                        break;

                    case TestOutcome.Inconclusive:
                        image = Properties.Resources.exclamation_diamond;
                        break;
                }
            }

            image.MakeTransparent(Color.Fuchsia);
            return image;
        }
    }
}