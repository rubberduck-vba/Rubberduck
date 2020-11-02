using System.Collections.Generic;
using System.Linq;
using Moq;
using NUnit.Framework;
using Rubberduck.Parsing.Settings;
using Rubberduck.Parsing.VBA.DeclarationResolving;
using Rubberduck.SettingsProvider;
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
            var unprotectedProjectIds = new[] { ("test", "path1"), ("ignoredTest", "path3") };
            var lockedProjectIds = new[] { ("lockedTest", "path2"), ("IrgnoredLockedTest", "path4") };
            var selector = TestSelector(unprotectedProjectIds, lockedProjectIds, "path3", "path4");

            var toBeResolvedFromCom = selector.ProjectsToResolveFromComProjects;

            Assert.False(toBeResolvedFromCom.Any());
        }

        [Test]
        public void AfterRefreshProjectToResolveFromComContainsLockedProjects()
        {
            var unprotectedProjectIds = new[] { ("test", "path1") };
            var lockedProjectIds = new[] { ("lockedTest", "path2") };
            var selector = TestSelector(unprotectedProjectIds, lockedProjectIds);

            selector.RefreshProjectsToResolveFromComProjectSelector();
            var toBeResolvedFromCom = selector.ProjectsToResolveFromComProjects;

            Assert.True(toBeResolvedFromCom.Contains("lockedTest"));
        }

        [Test]
        public void AfterRefreshProjectToResolveFromComContainsIgnoredProjects()
        {
            var unprotectedProjectIds = new[] { ("test", "path1"), ("ignoredTest", "path3") };
            var lockedProjectIds = Enumerable.Empty<(string projectId, string filename)>();
            var selector = TestSelector(unprotectedProjectIds, lockedProjectIds, "path3");

            selector.RefreshProjectsToResolveFromComProjectSelector();
            var toBeResolvedFromCom = selector.ProjectsToResolveFromComProjects;

            Assert.True(toBeResolvedFromCom.Contains("ignoredTest"));
        }

        [Test]
        public void AfterRemovalOfProjectWithoutRefreshProjectToResolveFromComStillContainsProject()
        {
            var unprotectedProjectIds = new[] { ("test", "path1") };
            var lockedProjectIds = new[] { ("lockedTest", "path2") };
            var mockProjectsProvider = MockTestProjectsProvider(unprotectedProjectIds, lockedProjectIds);
            var configService = ConfigService();
            var selector = new ProjectsToResolveFromComProjectsSelector(mockProjectsProvider.Object, configService);
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
            var unprotectedProjectIds = new[] { ("test", "path1") };
            var lockedProjectIds = new[] { ("lockedTest", "path2") };
            var mockProjectsProvider = MockTestProjectsProvider(unprotectedProjectIds, lockedProjectIds);
            var configService = ConfigService();
            var selector = new ProjectsToResolveFromComProjectsSelector(mockProjectsProvider.Object, configService);
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
            var unprotectedProjectIds = new[] { ("test", "path1")};
            var lockedProjectIds = new[] { ("lockedTest", "path2") };
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
            var unprotectedProjectIds = new[] { ("test", "path1"), ("ignoredTest", "path3") };
            var lockedProjectIds = new[] { ("lockedTest", "path2") };
            var selector = TestSelector(unprotectedProjectIds, lockedProjectIds, "path3");

            selector.RefreshProjectsToResolveFromComProjectSelector();
            var toBeResolvedFromCom = selector.ProjectsToResolveFromComProjects;
            if (!toBeResolvedFromCom.Contains("lockedTest") || !toBeResolvedFromCom.Contains("ignoredTest"))
            {
                Assert.Inconclusive("test projectIds not loaded");
            }

            Assert.True(selector.ToBeResolvedFromComProject("lockedTest"));
            Assert.True(selector.ToBeResolvedFromComProject("ignoredTest"));
        }


        private IProjectsToResolveFromComProjectSelector TestSelector(IEnumerable<(string projectId, string filename)> unprotectedProjectIds, IEnumerable<(string projectId, string filename)> lockedProjectIds, params string[] ignoredProjectFilenames)
        {
            var projectsProvider = MockTestProjectsProvider(unprotectedProjectIds, lockedProjectIds).Object;
            var configService = ConfigService(ignoredProjectFilenames);
            return new ProjectsToResolveFromComProjectsSelector(projectsProvider, configService);
        }

        private Mock<IProjectsProvider> MockTestProjectsProvider(IEnumerable<(string projectId, string filename)> unprotectedProjectIds, IEnumerable<(string projectId, string filename)> lockedProjectIds)
        {
            var mock = new Mock<IProjectsProvider>();
            mock.Setup(m => m.Projects()).Returns(() => unprotectedProjectIds.Select(tpl => (tpl.projectId, MockProject(tpl.projectId, tpl.filename))));
            mock.Setup(m => m.LockedProjects()).Returns(() => lockedProjectIds.Select(tpl => (tpl.projectId, MockProject(tpl.projectId, tpl.filename))));
            return mock;
        }

        private static IVBProject MockProject(string projectId, string filename)
        {
            var mock = new Mock<IVBProject>();

            mock.Setup(m => m.ProjectId).Returns(projectId);
            mock.Setup(m => m.FileName).Returns(filename);

            return mock.Object;
        }

        private static IConfigurationService<IgnoredProjectsSettings> ConfigService(params string[] ignoredProjectFilenames)
        {
            var settings = new IgnoredProjectsSettings
            {
                IgnoredProjectPaths = ignoredProjectFilenames.ToList()
            };

            var serviceMock = new Mock<IConfigurationService<IgnoredProjectsSettings>>();
            serviceMock.Setup(m => m.Read()).Returns(settings);

            return serviceMock.Object;
        }
    }
}