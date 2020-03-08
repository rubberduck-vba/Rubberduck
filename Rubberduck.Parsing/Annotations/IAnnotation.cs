using System;
using System.Collections.Generic;

namespace Rubberduck.Parsing.Annotations
{
    public interface IAnnotation
    {
        string Name { get; }

        /// <summary>
        /// The kind of object this annotation can be applied to 
        /// </summary>
        AnnotationTarget Target { get; }

        /// <summary>
        /// Specifies whether there can be multiple instances of the annotation on the same target.
        /// </summary>
        bool AllowMultiple { get; }

        /// <summary>
        /// The minimal number of arguments that must be provided to for this annotation
        /// </summary>
        int RequiredArguments { get; }

        /// <summary>
        /// The maximal number of arguments that must be provided to for this annotation; null means that there is no limit.
        /// </summary>
        int? AllowedArguments { get; }

        IReadOnlyList<string> ProcessAnnotationArguments(IEnumerable<string> arguments);
    }

    [Flags]
    public enum AnnotationTarget
    {
        /// <summary>
        /// Indicates that the annotation is valid for modules.
        /// </summary>
        Module = 1 << 0,
        /// <summary>
        /// Indicates that the annotation is valid for members.
        /// </summary>
        Member = 1 << 1,
        /// <summary>
        /// Indicates that the annotation is valid for variables or constants.
        /// </summary>
        Variable = 1 << 2,
        /// <summary>
        /// Indicates that the annotation is valid for identifier references.
        /// </summary>
        Identifier = 1 << 3,
        /// <summary>
        /// A convenience access indicating that the annotation is valid for Members, Variables and Identifiers.
        /// </summary>
        General = Member | Variable | Identifier,
    }
}
