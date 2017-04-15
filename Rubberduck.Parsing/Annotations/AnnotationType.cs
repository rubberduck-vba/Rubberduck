using System;

namespace Rubberduck.Parsing.Annotations
{
    public enum AnnotationType
    {
        TestModule = 1 << 0,
        ModuleInitialize = 1 << 1,
        ModuleCleanup = 1 << 2,
        TestMethod = 1 << 3,
        TestInitialize = 1 << 4,
        TestCleanup = 1 << 5,
        IgnoreTest = 1 << 6,
        Ignore = 1 << 7,
        IgnoreModule = 1 << 8,
        Folder = 1 << 9,
        NoIndent = 1 << 10,
        Interface = 1 << 11,
        /// <summary>
        /// Attribute annotations drive member attributes.
        /// </summary>
        Attribute = 1 << 12,
        [AttributeAnnotation("VB_Description")]
        Description = 1 << 13 | Attribute,
        [AttributeAnnotation("VB_UserMemId")]
        DefaultMember = 1 << 14 | Attribute,
        [AttributeAnnotation("VB_UserMemId")]
        Enumerator = 1 << 15 | Attribute,
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
