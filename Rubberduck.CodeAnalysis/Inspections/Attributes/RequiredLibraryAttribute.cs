using System;

namespace Rubberduck.CodeAnalysis.Inspections.Attributes
{
    /// <summary>
    /// This inspection requires a specific type library to be referenced in order to run.
    /// </summary>
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = true)]
    internal class RequiredLibraryAttribute : Attribute
    {
        public RequiredLibraryAttribute(string name)
        {
            LibraryName = name;
        }

        public string LibraryName { get; }
    }
}