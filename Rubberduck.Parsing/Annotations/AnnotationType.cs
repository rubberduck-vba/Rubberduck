using System;

namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// Member names are 
    /// </summary>
    public enum AnnotationType
    {
        /// <summary>
        /// A flag indicating that the annotation type is valid once per module.
        /// </summary>
        ModuleAnnotation = 1 << 1,
        /// <summary>
        /// A flag indicating that the annotation type is valid once per member.
        /// </summary>
        MemberAnnotation = 1 << 2,

        /// <summary>
        /// A flag indicating that the annotation type is driving an attribute.
        /// </summary>
        Attribute = 1 << 4,

        TestModule = 1 << 8 | ModuleAnnotation,
        ModuleInitialize = 1 << 9 | MemberAnnotation,
        ModuleCleanup = 1 << 10 | MemberAnnotation,
        TestMethod = 1 << 11 | MemberAnnotation,
        TestInitialize = 1 << 12 | MemberAnnotation,
        TestCleanup = 1 << 13 | MemberAnnotation,
        IgnoreTest = 1 << 14 | MemberAnnotation,
        Ignore = 1 << 15,
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
    }

    [AttributeUsage(AttributeTargets.Field)]
    public class AttributeAnnotationAttribute : Attribute
    {
        /// <summary>
        /// Enum value is associated with a VB_Attribute.
        /// </summary>
        /// <param name="name">The name of the associated attribute.</param>
        /// <param name="value">If specified, contrains the association to a specific value.</param>
        public AttributeAnnotationAttribute(string name, string value = null)
        {
            AttributeName = name;
            AttributeValue = value; // null default is assumed in AttributeExtensions.AnnotationType()
        }

        public string AttributeName { get; }
        public string AttributeValue { get; }
    }
}
