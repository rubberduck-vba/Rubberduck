using System;
using System.Collections.Generic;

namespace Rubberduck.Parsing.Annotations
{
    public interface IAttributeAnnotation : IAnnotation
    {
        string Attribute { get; }
        IReadOnlyList<string> AttributeValues { get; }
    }
    // attributes are disjoint to avoid issues around security and multiple attributes
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false, Inherited = false)]
    public class FixedAttributeValueAnnotationAttribute : Attribute
    {
        /// <summary>
        /// Enum value is associated with a VB_Attribute  with a fixed value.
        /// </summary>
        /// <param name="name">The name of the associated attribute.</param>
        /// <param name="value">If specified, constrains the association to a specific value.</param>
        public FixedAttributeValueAnnotationAttribute(string name, params string[] values)
        {
            AttributeName = name;
            AttributeValues = values;
        }

        public string AttributeName { get; }
        public IReadOnlyList<string> AttributeValues { get; }
    }

    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false, Inherited = false)]
    public class FlexibleAttributeValueAnnotationAttribute : Attribute
    {
        /// <summary>
        /// Enum value is associated with a VB_Attribute with a fixed number of values taken from the annotation values.
        /// </summary>
        /// <param name="name">The name of the associated attribute.</param>
        /// <param name="numberOfParameters">Size of the attribute value list the attribute takes.</param>
        /// <param name="toAnnotationValues">
        /// A function used during parsing to transform the values stored in the exported attribute to those stored in the code pass annotation arguments.
        /// </param>
        public FlexibleAttributeValueAnnotationAttribute(string name, int numberOfParameters, bool hasCustomTransform = false)
        {
            AttributeName = name;
            NumberOfParameters = numberOfParameters;
            HasCustomTransformation = hasCustomTransform;
        }

        public string AttributeName { get; }
        public int NumberOfParameters { get; }
        public bool HasCustomTransformation { get; }
    }
}