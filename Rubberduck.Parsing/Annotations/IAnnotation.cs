using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;
using System;

namespace Rubberduck.Parsing.Annotations
{
    public interface IAnnotation
    {
        QualifiedSelection QualifiedSelection { get; }
        VBAParser.AnnotationContext Context { get; }
        int? AnnotatedLine { get; }
        AnnotationAttribute MetaInformation { get; }

        string AnnotationType { get; }
    }

    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false, Inherited = false)]
    public class AnnotationAttribute : Attribute
    {
        public string Name { get; }
        public AnnotationTarget Target { get; }
        public bool AllowMultiple { get; }

        public AnnotationAttribute(string name, AnnotationTarget target, bool allowMultiple = false)
        {
            Name = name;
            Target = target;
            AllowMultiple = allowMultiple;
        }
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
