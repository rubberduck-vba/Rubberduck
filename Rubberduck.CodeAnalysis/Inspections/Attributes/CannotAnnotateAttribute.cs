using System;

namespace Rubberduck.CodeAnalysis.Inspections.Attributes
{
    /// <summary>
    /// This inspection isn't looking at code from the CodePane pass, and cannot be annotated.
    /// </summary>
    [AttributeUsage(AttributeTargets.Class)]
    internal class CannotAnnotateAttribute : Attribute
    {}
}