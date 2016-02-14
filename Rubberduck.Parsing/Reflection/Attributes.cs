namespace Rubberduck.Parsing.Reflection
{
    /// <summary>
    /// A base class for all member attributes.
    /// </summary>
    public abstract class MemberAttributeBase
    {
        public string Name { get { return GetType().Name.Replace("Attribute", string.Empty); } }
    }

    /// <summary>
    /// An attribute that marks a code Module as a test Module.
    /// </summary>
    public class TestModuleAttribute : MemberAttributeBase { }

    /// <summary>
    /// An attribute that marks a public procedure as a test method.
    /// </summary>
    public class TestMethodAttribute : MemberAttributeBase { }

    /// <summary>
    /// An attribute that marks a public procedure as a method to execute before each test is executed.
    /// </summary>
    public class TestInitializeAttribute : MemberAttributeBase { }

    /// <summary>
    /// An attribute that marks a public procedure as a method to execute after each test is executed.
    /// </summary>
    public class TestCleanupAttribute : MemberAttributeBase { }

    /// <summary>
    /// An attribute that marks a public procedure as a method to execute before the first test is executed.
    /// </summary>
    public class ModuleInitializeAttribute : MemberAttributeBase { }

    /// <summary>
    /// An attribute that marks a public procedure as a method to execute after all tests in the module have executed.
    /// </summary>
    public class ModuleCleanupAttribute : MemberAttributeBase { }
}
