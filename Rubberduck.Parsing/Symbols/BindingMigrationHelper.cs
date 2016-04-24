using Antlr4.Runtime;

namespace Rubberduck.Parsing.Symbols
{
    public static class BindingMigrationHelper
    {
        public static bool HasParent<T>(RuleContext context)
        {
            if (context == null)
            {
                return false;
            }
            if (context.Parent is T)
            {
                return true;
            }
            return HasParent<T>(context.Parent);
        }
    }
}
