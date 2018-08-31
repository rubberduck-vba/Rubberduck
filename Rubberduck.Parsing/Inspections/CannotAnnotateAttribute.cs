using System;

namespace Rubberduck.Parsing.Inspections
{
    /// <summary>
    /// This inspection isn't looking at code from the CodePane pass, and cannot be annotated.
    /// </summary>
    [AttributeUsage(AttributeTargets.Class)]
    public class CannotAnnotateAttribute : Attribute
    {
    }
}