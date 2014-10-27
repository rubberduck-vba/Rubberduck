using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RetailCoderVBE.Reflection
{
    /// <summary>
    /// A base class for all member attributes.
    /// </summary>
    internal abstract class MemberAttributeBase
    {
        public string Name { get { return GetType().Name.Replace("Attribute", string.Empty); } }
    }

    /// <summary>
    /// An attribute that marks a code module as a test module.
    /// </summary>
    internal class TestModuleAttribute : MemberAttributeBase { }

    /// <summary>
    /// An attribute that marks a public procedure as a test method.
    /// </summary>
    internal class TestMethodAttribute : MemberAttributeBase { }

    /// <summary>
    /// An attribute that marks a public procedure as a method to execute before each test is executed.
    /// </summary>
    internal class TestInitializeAttribute : MemberAttributeBase { }

    /// <summary>
    /// An attribute that marks a public procedure as a method to execute after each test is executed.
    /// </summary>
    internal class TestCleanupAttribute : MemberAttributeBase { }
}
