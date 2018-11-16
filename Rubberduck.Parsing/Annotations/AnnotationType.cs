using System;

namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// Member names are 
    /// </summary>
    public enum AnnotationType
    {
        /// <summary>
        /// A type for all not recognized annotations.
        /// </summary>
        NotRecognized = 0,

        /// <summary>
        /// A flag indicating that the annotation type is valid for modules.
        /// </summary>
        ModuleAnnotation = 1 << 1,

        /// <summary>
        /// A flag indicating that the annotation type is valid for members (method).
        /// </summary>
        MemberAnnotation = 1 << 2,

        /// <summary>
        /// A flag indicating that the annotation type is valid for variables or constants.
        /// </summary>
        VariableAnnotation = 1 << 3,

        /// <summary>
        /// A flag indicating that the annotation type is valid for identifier references.
        /// </summary>
        IdentifierAnnotation = 1 << 4,

        /// <summary>
        /// A flag indicating that the annotation type is valid on everything but modules.
        /// </summary>
        GeneralAnnotation = 1 << 5 | MemberAnnotation | VariableAnnotation | IdentifierAnnotation,

        /// <summary>
        /// A flag indicating that the annotation type is driving an attribute.
        /// </summary>
        Attribute = 1 << 6,

        TestModule = 1 << 8 | ModuleAnnotation,
        ModuleInitialize = 1 << 9 | MemberAnnotation,
        ModuleCleanup = 1 << 10 | MemberAnnotation,
        TestMethod = 1 << 11 | MemberAnnotation,
        TestInitialize = 1 << 12 | MemberAnnotation,
        TestCleanup = 1 << 13 | MemberAnnotation,
        IgnoreTest = 1 << 14 | MemberAnnotation,
        Ignore = 1 << 15 | GeneralAnnotation,
        IgnoreModule = 1 << 16 | ModuleAnnotation,
        Folder = 1 << 17 | ModuleAnnotation,
        NoIndent = 1 << 18 | ModuleAnnotation,
        Interface = 1 << 19 | ModuleAnnotation,
        [AttributeAnnotation("VB_Description")]
        Description = 1 << 13 | Attribute | MemberAnnotation,
        [AttributeAnnotation("VB_UserMemId", "0")]
        DefaultMember = 1 << 14 | Attribute | MemberAnnotation,
        [AttributeAnnotation("VB_UserMemId", "-4")]
        Enumerator = 1 << 15 | Attribute | MemberAnnotation,
        [AttributeAnnotation("VB_PredeclaredId", "True")]
        PredeclaredId = 1 << 16 | Attribute | ModuleAnnotation,
        [AttributeAnnotation("VB_Exposed", "True")]
        Exposed = 1 << 17 | Attribute | ModuleAnnotation,
        Obsolete = 1 << 18 | MemberAnnotation | VariableAnnotation
    }

    [AttributeUsage(AttributeTargets.Field)]
    public class AttributeAnnotationAttribute : Attribute
    {
        /// <summary>
        /// Enum value is associated with a VB_Attribute.
        /// </summary>
        /// <param name="name">The name of the associated attribute.</param>
        /// <param name="value">If specified, constrains the association to a specific value.</param>
        public AttributeAnnotationAttribute(string name, string value = null)
        {
            AttributeName = name;
            AttributeValue = value; // null default is assumed in AttributeExtensions.AnnotationType()
        }

        public string AttributeName { get; }
        public string AttributeValue { get; }
    }
}
