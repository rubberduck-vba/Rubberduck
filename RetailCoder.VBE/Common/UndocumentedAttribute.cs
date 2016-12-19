using System;

namespace Rubberduck.Common
{
    /// <summary>
    /// Mark a feature as undocumented.
    /// </summary>
    /// <remarks>The RubberduckWeb project may this attribute to filter viewable content.</remarks>
    [AttributeUsage(AttributeTargets.Class)]
    public class UndocumentedAttribute : Attribute
    {
    }
}