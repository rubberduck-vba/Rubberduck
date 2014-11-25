using System.Drawing;
using System.Runtime.InteropServices;
using Rubberduck.UnitTesting;

namespace Rubberduck.UI.UnitTesting
{
    [ComVisible(false)]
    public static class TestResultExtensions
    {
        public static Image Icon(this TestResult result)
        {
            var image = Properties.Resources.Serious;
            if (result != null)
            {
                switch (result.Outcome)
                {
                    case TestOutcome.Succeeded:
                        image = Properties.Resources.OK;
                        break;

                    case TestOutcome.Failed:
                        image = Properties.Resources.Critical;
                        break;

                    case TestOutcome.Inconclusive:
                        image = Properties.Resources.Warning;
                        break;
                }
            }

            image.MakeTransparent(Color.Fuchsia);
            return image;
        }
    }
}