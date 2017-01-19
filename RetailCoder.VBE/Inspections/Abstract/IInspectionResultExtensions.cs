using System.Collections.Generic;
using System.Linq;
using Rubberduck.Properties;

namespace Rubberduck.Inspections.Abstract
{
    public static class IInspectionResultExtensions
    {
        public static IFixableResult AsFixable([NotNull] this Parsing.Symbols.IInspectionResult source)
        {
            return source as IFixableResult;
        }

        public static IEnumerable<IFixableResult> AsFixable([NotNull] this IEnumerable<Parsing.Symbols.IInspectionResult> source)
        {
            return source.OfType<IFixableResult>();
        }
    }
}