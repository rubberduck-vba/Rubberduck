using System;

namespace Rubberduck.Parsing.Annotations
{
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
        ModuleInitialize = 1 << 9 | ModuleAnnotation,
        ModuleCleanup = 1 << 10 | ModuleAnnotation,
        TestMethod = 1 << 11 | MemberAnnotation,
        TestInitialize = 1 << 12 | ModuleAnnotation,
        TestCleanup = 1 << 13 | ModuleAnnotation,
        IgnoreTest = 1 << 14 | MemberAnnotation,
        Ignore = 1 << 15,
        IgnoreModule = 1 << 16 | ModuleAnnotation,
        Folder = 1 << 17 | ModuleAnnotation,
        NoIndent = 1 << 18 | ModuleAnnotation,
        Interface = 1 << 19 | ModuleAnnotation,
        [AttributeAnnotation("VB_Description")]
        Description = 1 << 13 | Attribute | MemberAnnotation,
        [AttributeAnnotation("VB_UserMemId")]
        DefaultMember = 1 << 14 | Attribute | ModuleAnnotation,
        [AttributeAnnotation("VB_UserMemId")]
        Enumerator = 1 << 15 | Attribute | MemberAnnotation,
        [AttributeAnnotation("VB_PredeclaredId")]
        PredeclaredId = 1 << 16 | Attribute | ModuleAnnotation,
        [AttributeAnnotation("VB_Exposed")]
        Exposed = 1 << 17 | Attribute | ModuleAnnotation,
    }

    [AttributeUsage(AttributeTargets.Field)]
    public class AttributeAnnotationAttribute : Attribute
    {
        public AttributeAnnotationAttribute(string name)
        {
            AttributeName = name;
        }

        public string AttributeName { get; }
    }
}
