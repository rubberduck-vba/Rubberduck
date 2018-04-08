using System;

namespace Rubberduck.Parsing.Common
{
    [AttributeUsage(AttributeTargets.Class)]
    public class ExperimentalAttribute : Attribute
    {
        public ExperimentalAttribute() : this(string.Empty)
        {
        }

        public ExperimentalAttribute(string resource)
        {
        }
    }
}