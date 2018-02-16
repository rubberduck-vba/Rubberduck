using NUnit.Framework;
using Rubberduck.VBEditor.ComManagement;

namespace RubberduckTests.VBEditor
{
    [TestFixture()]
    public class StrongComSafeTests : ComSafeTestBase
    {
        protected override IComSafe TestComSafe()
        {
            return new StrongComSafe();
        }
    }

    [TestFixture()]
    public class WeakComSafeTests : ComSafeTestBase
    {
        protected override IComSafe TestComSafe()
        {
            return new WeakComSafe();
        }
    }
}
