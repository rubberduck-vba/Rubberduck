using System.Drawing;
using Rubberduck.Properties;
using Rubberduck.UnitTesting;

namespace Rubberduck.UI.UnitTesting
{
    public static class TestResultExtensions
    {
        public static Image Icon(this TestResult result)
        {
            var image = Resources.question_white;
            if (result != null)
            {
                switch (result.Outcome)
                {
                    case TestOutcome.Succeeded:
                        image = Resources.tick_circle;
                        break;

                    case TestOutcome.Failed:
                        image = Resources.cross_circle;
                        break;

                    case TestOutcome.Inconclusive:
                        image = Resources.exclamation_diamond;
                        break;
                }
            }

            image.MakeTransparent(Color.Fuchsia);
            return image;
        }
    }
}