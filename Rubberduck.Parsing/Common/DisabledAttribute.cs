using System;
using System.Diagnostics;

namespace Rubberduck.Parsing.Common
{ 
// Conditional cannot be expressed negatively, so we use combined #if and Conditional
// to ensure that the code is always checked but omitted in non-debug builds.
#if !DEBUG 
    [Conditional("I_WILL_NEVER_EXIST_CRY_MY_BELOVED_DUCK")]
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
