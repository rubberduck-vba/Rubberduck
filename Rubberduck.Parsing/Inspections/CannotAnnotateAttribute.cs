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

    [AttributeUsage(AttributeTargets.Class, AllowMultiple = true)]
    public class QuickFixAttribute : Attribute
    {
        public QuickFixAttribute(Type quickFixType)
        {
            QuickFixType = quickFixType;
        }

        public Type QuickFixType { get; }
    }
}