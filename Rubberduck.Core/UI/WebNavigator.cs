using System;
using System.Diagnostics;

namespace Rubberduck.UI
{
    public interface IWebNavigator
    {
        /// <summary>
        /// Opens the specified URI in the default browser.
        /// </summary>
        void Navigate(Uri uri);
    }

    public class WebNavigator : IWebNavigator
    {
        public void Navigate(Uri uri)
        {
            Process.Start(new ProcessStartInfo(uri.AbsoluteUri));
        }
    }
}
