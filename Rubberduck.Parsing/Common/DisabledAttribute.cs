using System;

namespace Rubberduck.Parsing.Common
{
    // Conditional cannot be an expression, so we use combined #if and an
    // non-existent flag to ensure that the code is always checked at the
    // compile time but omitted in non-debug builds.
#if !DEBUG
    [System.Diagnostics.Conditional("I_WILL_NEVER_EXIST_CRY_MY_BELOVED_DUCK")]
#endif
    [AttributeUsage(AttributeTargets.Class)]
    public class DisabledAttribute : Attribute
    {
        public DisabledAttribute() : this(string.Empty)
        {
        }

        public DisabledAttribute(string resource)
        {
        }
    }
}
