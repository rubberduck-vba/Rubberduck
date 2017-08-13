using System;

namespace Rubberduck.Parsing.Inspections
{
    /// <summary>
    /// This inspection requires a specific type library to be referenced in order to run.
    /// </summary>
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