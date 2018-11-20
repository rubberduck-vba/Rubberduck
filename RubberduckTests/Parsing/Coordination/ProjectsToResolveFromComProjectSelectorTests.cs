using System.Collections.Generic;
using System.Linq;
using Moq;
using NUnit.Framework;
using Rubberduck.Parsing.VBA.DeclarationResolving;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace RubberduckTests.Parsing.Coordination
{
    [TestFixture]
    public class ProjectsToResolveFromComProjectSelectorTests
    {
        [Test]
        public void InitiallyNoProjectToResolveFromCom()
        {
            var unprotectedProjectIds = new[] {"test"};
            var lockedProjectIds = new[] { "lockedTest" };
            var selector = TestSelector(unprotectedProjectIds, lockedProjectIds);

            var toBeResolvedFromCom = selector.ProjectsToResolveFromComProjects;

            Assert.False(toBeResolvedFromCom.Any());
        }


        [Test]
        public void AfterRefreshProjectToResolveFromComContainsLockedProjects()
        {
            var unprotectedProjectIds = new[] { "test" };
            var lockedProjectIds = new[] { "lockedTest" };
            var selector = TestSelector(unprotectedProjectIds, lockedProjectIds);

            selector.RefreshProjectsToResolveFromComProjectSelector();
            var toBeResolvedFromCom = selector.ProjectsToResolveFromComProjects;

            Assert.True(toBeResolvedFromCom.Contains("lockedTest"));
        }


        [Test]
        public void AfterRemovalOfProjectWithoutRefreshProjectToResolveFromComStillContainsProject()
        {
            var unprotectedProjectIds = new[] { "test" };
            var lockedProjectIds = new[] { "lockedTest" };
            var mockProjectsProvider = MockTestProjectsProvider(unprotectedProjectIds, lockedProjectIds);
            var selector = new ProjectsToResolveFromComProjectsSelector(mockProjectsProvider.Object);
            selector.RefreshProjectsToResolveFromComProjectSelector();

            var beforeRemoval = selector.ProjectsToResolveFromComProjects;
            if (!beforeRemoval.Contains("lockedTest"))
            {
                Assert.Inconclusive("test projectId to remove has not been loaded");
            }

            mockProjectsProvider.Setup(m => m.LockedProjects()).Returns(() => new List<(string ProjectId, IVBProject Project)>());

            var toBeResolvedFromCom = selector.ProjectsToResolveFromComProjects;

            Assert.True(toBeResolvedFromCom.Contains("lockedTest"));
        }


        [Test]
        public void AfterRemovalOfProjectWithRefreshProjectToResolveFromComNoLongerContainsProject()
        {
            var unprotectedProjectIds = new[] { "test" };
            var lockedProjectIds = new[] { "lockedTest" };
            var mockProjectsProvider = MockTestProjectsProvider(unprotectedProjectIds, lockedProjectIds);
            var selector = new ProjectsToResolveFromComProjectsSelector(mockProjectsProvider.Object);
            selector.RefreshProjectsToResolveFromComProjectSelector();

            var beforeRemoval = selector.ProjectsToResolveFromComProjects;
            if (!beforeRemoval.Contains("lockedTest"))
            {
                Assert.Inconclusive("test projectId to remove has not been loaded");
            }

            mockProjectsProvider.Setup(m => m.LockedProjects()).Returns(() => new List<(string ProjectId, IVBProject Project)>());
            selector.RefreshProjectsToResolveFromComProjectSelector();

            var toBeResolvedFromCom = selector.ProjectsToResolveFromComProjects;

            Assert.False(toBeResolvedFromCom.Any());
        }

        [Test]
        public void ToBeResolvedFromComReturnsFalseForProjectIdsNotInToBeResolvedFromCom()
        {
            var unprotectedProjectIds = new[] { "test" };
            var lockedProjectIds = new[] { "lockedTest" };
            var selector = TestSelector(unprotectedProjectIds, lockedProjectIds);

            var toBeResolvedFromCom = selector.ProjectsToResolveFromComProjects;
            if (toBeResolvedFromCom.Contains("lockedTest"))
            {
                Assert.Inconclusive("test projectId already loaded");
            }

            Assert.False(selector.ToBeResolvedFromComProject("lockedTest"));
        }

        [Test]
        public void ToBeResolvedFromComReturnsTrueForProjectIdsInToBeResolvedFromCom()
        {
            var unprotectedProjectIds = new[] { "test" };
            var lockedProjectIds = new[] { "lockedTest" };
            var selector = TestSelector(unprotectedProjectIds, lockedProjectIds);

            selector.RefreshProjectsToResolveFromComProjectSelector();
            var toBeResolvedFromCom = selector.ProjectsToResolveFromComProjects;
            if (!toBeResolvedFromCom.Contains("lockedTest"))
            {
                Assert.Inconclusive("test projectId not loaded");
            }

            Assert.True(selector.ToBeResolvedFromComProject("lockedTest"));
        }


        private IProjectsToResolveFromComProjectSelector TestSelector(IEnumerable<string> unprotectedProjectIds, IEnumerable<string> lockedProjectIds)
        {
            var projectsProvider = MockTestProjectsProvider(unprotectedProjectIds, lockedProjectIds).Object;
            return new ProjectsToResolveFromComProjectsSelector(projectsProvider);
        }

        private Mock<IProjectsProvider> MockTestProjectsProvider(IEnumerable<string> unprotectedProjectIds, IEnumerable<string> lockedProjectIds)
        {
            var mock = new Mock<IProjectsProvider>();
            mock.Setup(m => m.Projects()).Returns(() => unprotectedProjectIds.Select(projectId => (projectId, new Mock<IVBProject>().Object)));
            mock.Setup(m => m.LockedProjects()).Returns(() => lockedProjectIds.Select(projectId => (projectId, new Mock<IVBProject>().Object)));
            return mock;
        }
    }
}