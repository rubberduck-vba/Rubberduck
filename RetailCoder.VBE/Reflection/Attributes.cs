using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Reflection
{
    /// <summary>
    /// A base class for all member attributes.
    /// </summary>
    [ComVisible(false)]
    public abstract class MemberAttributeBase
    {
        public string Name { get { return GetType().Name.Replace("Attribute", string.Empty); } }
    }

    /// <summary>
    /// An attribute that marks a code Module as a test Module.
    /// </summary>
    [ComVisible(false)]
    public class TestModuleAttribute : MemberAttributeBase { }

    /// <summary>
    /// An attribute that marks a public procedure as a test method.
    /// </summary>
    [ComVisible(false)]
    public class TestMethodAttribute : MemberAttributeBase { }

    /// <summary>
    /// An attribute that marks a public procedure as a method to execute before each test is executed.
    /// </summary>
    [ComVisible(false)]
    public class TestInitializeAttribute : MemberAttributeBase { }

    /// <summary>
    /// An attribute that marks a public procedure as a method to execute after each test is executed.
    /// </summary>
    [ComVisible(false)]
    public class TestCleanupAttribute : MemberAttributeBase { }
}
