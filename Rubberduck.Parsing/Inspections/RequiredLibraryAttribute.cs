using System;

namespace Rubberduck.Parsing.Inspections
{
    [AttributeUsage(AttributeTargets.Class)]
    public class RequiredLibraryAttribute : Attribute
    {
        public RequiredLibraryAttribute(string name)
        {
            LibraryName = name;
        }

        public string LibraryName { get; }
    }
}