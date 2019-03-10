using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Resources;

namespace Rubberduck.Resources
{
    public static class Locales
    {
        private static List<CultureInfo> _cultures;

        public static CultureInfo DefaultCulture => CultureInfo.GetCultureInfo("en-US");

        //Adapted from https://stackoverflow.com/a/32161480/4088852
        public static List<CultureInfo> AvailableCultures
        {
            get
            {
                if (!(_cultures is null))
                {
                    return _cultures;
                }

                _cultures = new List<CultureInfo> { DefaultCulture };
                var resources = new ResourceManager("Rubberduck.Resources.RubberduckUI", Assembly.GetAssembly(typeof(Locales)));
                foreach (var culture in CultureInfo.GetCultures(CultureTypes.AllCultures).Where(locale => !locale.Equals(CultureInfo.InvariantCulture)))
                {
                    try
                    {
                        if (!(resources.GetResourceSet(culture, true, false) is null))
                        {
                            _cultures.Add(culture);
                        }
                    }
                    catch (CultureNotFoundException)
                    {
                        // Ignored.
                    }
                }

                return _cultures;
            }
        }
    }
}
