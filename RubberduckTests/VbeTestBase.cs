using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using MockFactory = RubberduckTests.Mocks.MockFactory;

namespace RubberduckTests
{
    public abstract class VbeTestBase
    {
        private Mock<IVBE> _ide;
        private ICollection<IVBProject> _projects;

        [TestInitialize]
        public void Initialize()
        {
            _ide = MockFactory.CreateVbeMock();

            _projects = new List<IVBProject>();
            var projects = MockFactory.CreateProjectsMock(_projects);
            projects.Setup(m => m[It.IsAny<int>()]).Returns<int>(i => _projects.ElementAt(i));

            _ide.SetupGet(m => m.VBProjects).Returns(() => projects.Object);
        }

        [TestCleanup]
        public void Cleanup()
        {
            _ide = null;
        }
    }
}
