using Rubberduck.VBEditor.SafeComWrappers;
using System;
using System.Collections.Generic;

namespace Rubberduck.Parsing.Annotations
{
    public interface IAnnotation
    {
        /// <summary>
        /// The name of the annotation (without the @ prefix).
        /// </summary>
        string Name { get; }

        /// <summary>
        /// The kind of object this annotation can be applied to.
        /// </summary>
        AnnotationTarget Target { get; }

        /// <summary>
        /// The types of component that are incompatible with this annotation, if any.
        /// </summary>
        IReadOnlyList<ComponentType> IncompatibleComponentTypes { get; }

        /// <summary>
        /// If supplied, annotation is only valid for this component type.
        /// </summary>
        ComponentType? RequiredComponentType { get; }

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

        /// <summary>
        /// The allowed types of arguments for the annotation. The last argument type is valid for all optional arguments.
        /// </summary>
        IReadOnlyList<AnnotationArgumentType> AllowedArgumentTypes { get; }

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

    [Flags]
    public enum AnnotationArgumentType
    {
        /// <summary>
        /// Indicates that the annotation argument can be a string.
        /// </summary>
        Text = 1,
        /// <summary>
        /// Indicates that the annotation argument can be a number.
        /// </summary>
        Number = 1 << 1,
        /// <summary>
        /// Indicates that the annotation argument can be a boolean.
        /// </summary>
        Boolean = 1 << 2,
        /// <summary>
        /// Indicates that the annotation argument can be an inspection name.
        /// </summary>
        Inspection = 1 << 3,
        /// <summary>
        /// Indicates that the annotation argument can be an attribute name.
        /// </summary>
        Attribute = 1 << 4
    }
}
