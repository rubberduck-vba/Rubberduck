using System.Windows;
using System.Windows.Media;

namespace Rubberduck.UI.Controls
{
    public static class DependencyObjectExtensions
    {
        //from https://stackoverflow.com/a/41985834/1188513
        public static T GetAncestor<T>(this DependencyObject child, int maxLevels = 10) where T : DependencyObject
        {
            var levels = 0;
            var parent = child;
            do
            {
                parent = VisualTreeHelper.GetParent(parent);
                if (parent is T ancestor)
                {
                    return ancestor;
                }

                levels++;
            }
            while (parent != null && levels <= maxLevels);
            return null;
        }

    }
}