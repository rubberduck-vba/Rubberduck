using System;

namespace Rubberduck.Parsing.Common
{
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
