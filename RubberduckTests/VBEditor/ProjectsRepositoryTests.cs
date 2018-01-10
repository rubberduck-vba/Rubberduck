using System;
using System.Linq;
using Moq;
using NUnit.Framework;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.VBEditor
{
    [TestFixture()]
    public class ProjectsRepositoryTests
    {
        [Test()]
        public void ProjectsCollectionReturnsTheOneFromTheVbePassedIn()
        {
            var vbe = new MockVbeBuilder().Build().Object;
            var repository =  new ProjectsRepository(vbe);

            var vbePojects = vbe.VBProjects;
            var repositoryProjects = repository.ProjectsCollection();

            Assert.AreEqual(vbePojects, repositoryProjects);
        }

        [Test()]
        public void ProjectsCollectionGetsDisposedOnDisposal()
        {
            var vbeBuilder = new MockVbeBuilder();
            var vbe = vbeBuilder.Build().Object;
            var repository = new ProjectsRepository(vbe);

            repository.Dispose();

            var projectsMock = vbeBuilder.MockProjectsCollection;
            projectsMock.Verify(m => m.Dispose(), Times.Once);
        }

        [Test()]
        public void ProjectsCollectionReturnsNullAfterDisposal()
        {
            var vbe = new MockVbeBuilder().Build().Object;
            var repository = new ProjectsRepository(vbe);

            repository.Dispose();

            Assert.IsNull(repository.ProjectsCollection());
        }

        [Test()]
        public void ProjectsReturnsProjectsOnVbe()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            var projectMock = projectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            var otherProjectMock = otherProjectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var otherProject = otherProjectMock.Object;

            var repository = new ProjectsRepository(vbe);
            var projects = repository.Projects().ToList();

            Assert.AreEqual(2, projects.Count);
            Assert.Contains((project.ProjectId, project), projects);
            Assert.Contains((otherProject.ProjectId, otherProject), projects);
        }

        [Test()]
        public void ProjectsGetDisposedOnDisposal()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            var projectMock = projectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            var otherProjectMock = otherProjectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;

            var repository = new ProjectsRepository(vbe);
            repository.Dispose();

            projectMock.Verify(m => m.Dispose(), Times.Once);
            otherProjectMock.Verify(m => m.Dispose(), Times.Once);
        }

        [Test()]
        public void ProjectsReturnsEmptyCollectionAfterDisposal()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            var projectMock = projectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            var otherProjectMock = otherProjectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;

            var repository = new ProjectsRepository(vbe);
            repository.Dispose();
            var projects = repository.Projects().ToList();

            Assert.IsEmpty(projects);
        }

        [Test()]
        public void ProjectsDoesNotReturnProjectsAddedToVbeWithoutRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            var projectMock = projectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(projectMock);

            var vbe = vbeBuilder.Build().Object;

            var repository = new ProjectsRepository(vbe);
            var otherProject = vbe.VBProjects.Add(ProjectType.HostProject);
            var projects = repository.Projects().ToList();

            Assert.IsFalse(projects.Contains((otherProject.ProjectId, otherProject)));
        }

        [Test()]
        public void ProjectsReturnsProjectsAddedToVbeAfterRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            var projectMock = projectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(projectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;

            var repository = new ProjectsRepository(vbe);
            var otherProject = vbe.VBProjects.Add(ProjectType.HostProject);
            repository.Refresh();
            var projects = repository.Projects().ToList();

            Assert.AreEqual(2, projects.Count);
            Assert.Contains((project.ProjectId, project), projects);
            Assert.Contains((otherProject.ProjectId, otherProject), projects);
        }

        [Test()]
        public void ProjectsReturnsRemovedProjectsBeforeRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            var projectMock = projectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            var otherProjectMock = otherProjectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var otherProject = otherProjectMock.Object;

            var repository = new ProjectsRepository(vbe);
            vbe.VBProjects.Remove(otherProject);
            var projects = repository.Projects().ToList();

            Assert.AreEqual(2, projects.Count);
            Assert.Contains((project.ProjectId, project), projects);
            Assert.Contains((otherProject.ProjectId, otherProject), projects);
        }

        [Test()]
        public void ProjectsDoesNotReturnRemovedProjectsAfterRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            var projectMock = projectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            var otherProjectMock = otherProjectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var otherProject = otherProjectMock.Object;

            var repository = new ProjectsRepository(vbe);
            vbe.VBProjects.Remove(otherProject);
            repository.Refresh();
            var projects = repository.Projects().ToList();

            Assert.IsFalse(projects.Contains((otherProject.ProjectId, otherProject)));
        }

        [Test()]
        public void ProjectsGetDisposedOnRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            var projectMock = projectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            var otherProjectMock = otherProjectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;

            var repository = new ProjectsRepository(vbe);
            repository.Refresh();

            projectMock.Verify(m => m.Dispose(), Times.Once);
            otherProjectMock.Verify(m => m.Dispose(), Times.Once);
        }

        [Test()]
        public void RemovedProjectsGetDisposedOnRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            var projectMock = projectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            var otherProjectMock = otherProjectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var otherProject = otherProjectMock.Object;

            var repository = new ProjectsRepository(vbe);
            vbe.VBProjects.Remove(otherProject);
            repository.Refresh();

            otherProjectMock.Verify(m => m.Dispose(), Times.Once);
        }

        [Test()]
        public void ProjectReturnsProjectWithMatchingProjectId()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            var projectMock = projectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            var otherProjectMock = otherProjectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var otherProject = otherProjectMock.Object;

            var repository = new ProjectsRepository(vbe);
            var returnedProject = repository.Project(otherProject.ProjectId);

            Assert.Equals(otherProject, returnedProject);
        }

        [Test()]
        public void ProjectReturnsNullForUnknownProjectId()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            var projectMock = projectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            var otherProjectMock = otherProjectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;

            var repository = new ProjectsRepository(vbe);
            var returnedProject = repository.Project(new Guid().ToString());

            Assert.IsNull(returnedProject);
        }

        [Test()]
        public void ProjectReturnsNullAfterDisposal()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            var projectMock = projectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            var otherProjectMock = otherProjectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var otherProject = otherProjectMock.Object;

            var repository = new ProjectsRepository(vbe);
            var returnedProject = repository.Project(otherProject.ProjectId);

            Assert.IsNull(returnedProject);
        }

        [Test()]
        public void ProjectReturnsNullForProjectIdOfAddedProjectBeforeRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            var projectMock = projectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(projectMock);

            var vbe = vbeBuilder.Build().Object;

            var repository = new ProjectsRepository(vbe);
            var otherProject = vbe.VBProjects.Add(ProjectType.HostProject);
            var returnedProject = repository.Project(otherProject.ProjectId);

            Assert.IsNull(returnedProject);
        }

        [Test()]
        public void ProjectReturnsAddedProjectWithMatchingProjectIdAfterRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            var projectMock = projectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(projectMock);

            var vbe = vbeBuilder.Build().Object;

            var repository = new ProjectsRepository(vbe);
            var otherProject = vbe.VBProjects.Add(ProjectType.HostProject);
            repository.Refresh();
            var returnedProject = repository.Project(otherProject.ProjectId);

            Assert.Equals(otherProject, returnedProject);
        }

        [Test()]
        public void ProjectReturnsRemovedProjectWithMatchingProjectIdBeforeRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            var projectMock = projectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            var otherProjectMock = otherProjectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var otherProject = otherProjectMock.Object;

            var repository = new ProjectsRepository(vbe);
            vbe.VBProjects.Remove(otherProject);
            var returnedProject = repository.Project(otherProject.ProjectId);

            Assert.Equals(otherProject, returnedProject);
        }

        [Test()]
        public void ProjectReturnsNullForProjectIdOfRemovedProjectAfterRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            var projectMock = projectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            var otherProjectMock = otherProjectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var otherProject = otherProjectMock.Object;

            var repository = new ProjectsRepository(vbe);
            vbe.VBProjects.Remove(otherProject);
            repository.Refresh();
            var returnedProject = repository.Project(otherProject.ProjectId);

            Assert.IsNull(returnedProject);
        }

        [Test()]
        public void ComponentCollectionsOfProjectsGetDisposedOnDisposal()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            var projectMock = projectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(projectMock);
            var componentsCollectionMock = projectBuilder.MockVBComponents;
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            var otherProjectMock = otherProjectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(otherProjectMock);
            var otherComponentsCollectionMock = otherProjectBuilder.MockVBComponents;

            var vbe = vbeBuilder.Build().Object;

            var repository = new ProjectsRepository(vbe);
            repository.Dispose();

            componentsCollectionMock.Verify(m => m.Dispose(), Times.Once);
            otherComponentsCollectionMock.Verify(m => m.Dispose(), Times.Once);
        }

        [Test()]
        public void ComponentsCollectionsOfProjectsGetDisposedOnRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            var projectMock = projectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(projectMock);
            var componentsCollectionMock = projectBuilder.MockVBComponents;
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            var otherProjectMock = otherProjectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(otherProjectMock);
            var otherComponentsCollectionMock = otherProjectBuilder.MockVBComponents;

            var vbe = vbeBuilder.Build().Object;

            var repository = new ProjectsRepository(vbe);
            repository.Refresh();

            componentsCollectionMock.Verify(m => m.Dispose(), Times.Once);
            otherComponentsCollectionMock.Verify(m => m.Dispose(), Times.Once);
        }

        [Test()]
        public void ComponentsCollectionsOfRemovedProjectsGetDisposedOnRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            var projectMock = projectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(projectMock);
            var componentsCollectionMock = projectBuilder.MockVBComponents;
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            var otherProjectMock = otherProjectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var otherProject = otherProjectMock.Object;

            var repository = new ProjectsRepository(vbe);
            vbe.VBProjects.Remove(otherProject);
            repository.Refresh();

            componentsCollectionMock.Verify(m => m.Dispose(), Times.Once);
        }

        [Test()]
        public void ComponentsCollectionReturnsComponentsCollectionOfProjectWithMatchingProjectId()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            var projectMock = projectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            var otherProjectMock = otherProjectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;

            var repository = new ProjectsRepository(vbe);
            var returnedCollection = repository.ComponentsCollection(project.ProjectId);

            Assert.Equals(project.VBComponents, returnedCollection);
        }

        [Test()]
        public void ComponentsCollectionReturnsNullForUnknownProjectId()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            var projectMock = projectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            var otherProjectMock = otherProjectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;

            var repository = new ProjectsRepository(vbe);
            var returnedCollection = repository.ComponentsCollection(new Guid().ToString());

            Assert.IsNull(returnedCollection);
        }

        [Test()]
        public void ComponentsCollectionReturnsNullAfterDisposal()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            var projectMock = projectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            var otherProjectMock = otherProjectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;

            var repository = new ProjectsRepository(vbe);
            var returnedCollection = repository.ComponentsCollection(project.ProjectId);

            Assert.IsNull(returnedCollection);
        }

        [Test()]
        public void ComponentsCollectionReturnsNullForProjectIdOfAddedProjectBeforeRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            var projectMock = projectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(projectMock);

            var vbe = vbeBuilder.Build().Object;

            var repository = new ProjectsRepository(vbe);
            var otherProject = vbe.VBProjects.Add(ProjectType.HostProject);
            var returnedCollection = repository.ComponentsCollection(otherProject.ProjectId);

            Assert.IsNull(returnedCollection);
        }

        [Test()]
        public void ComponentCollectionReturnsComponentsCollectionOfAddedProjectWithMatchingProjectIdAfterRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            var projectMock = projectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(projectMock);

            var vbe = vbeBuilder.Build().Object;

            var repository = new ProjectsRepository(vbe);
            var otherProject = vbe.VBProjects.Add(ProjectType.HostProject);
            repository.Refresh();
            var returnedCollection = repository.ComponentsCollection(otherProject.ProjectId);

            Assert.Equals(otherProject.VBComponents, returnedCollection);
        }

        [Test()]
        public void ComponentsCollectionReturnsComponentsCollectionOfRemovedProjectWithMatchingProjectIdBeforeRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            var projectMock = projectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            var otherProjectMock = otherProjectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var otherProject = otherProjectMock.Object;

            var repository = new ProjectsRepository(vbe);
            vbe.VBProjects.Remove(otherProject);
            var returnedCollection = repository.ComponentsCollection(otherProject.ProjectId);

            Assert.Equals(otherProject.VBComponents, returnedCollection);
        }

        [Test()]
        public void ComponentsCollectionReturnsNullForProjectIdOfRemovedProjectAfterRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            var projectMock = projectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            var otherProjectMock = otherProjectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var otherProject = otherProjectMock.Object;

            var repository = new ProjectsRepository(vbe);
            vbe.VBProjects.Remove(otherProject);
            repository.Refresh();
            var returnedCollection = repository.ComponentsCollection(otherProject.ProjectId);

            Assert.IsNull(returnedCollection);
        }

        [Test()]
        public void ComponentCollectionOfProjectWithMatchingProjectIdGetsDisposedOnRefreshForProjectId()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            var projectMock = projectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(projectMock);
            var componentsCollectionMock = projectBuilder.MockVBComponents;
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            var otherProjectMock = otherProjectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;

            var repository = new ProjectsRepository(vbe);
            repository.Refresh(project.ProjectId);

            componentsCollectionMock.Verify(m => m.Dispose(), Times.Once);
        }

        [Test()]
        public void ComponentCollectionsOfProjectsWithNonMatchingProjectIdDoNotGetDisposedOnRefreshForProjectId()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            var projectMock = projectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            var otherProjectMock = otherProjectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(otherProjectMock);
            var otherComponentsCollectionMock = otherProjectBuilder.MockVBComponents;

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;

            var repository = new ProjectsRepository(vbe);
            repository.Refresh(project.ProjectId);

            otherComponentsCollectionMock.Verify(m => m.Dispose(), Times.Never);
        }

        [Test()]
        public void ComponentsReturnsComponentsOnVbeWithQmns()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            projectBuilder.AddComponent("component2", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.BuildWIthAssignedProjectId();
            var mockComponents = projectBuilder.MockComponents;
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(otherProjectMock);
            var otherMockComponents = otherProjectBuilder.MockComponents;

            var vbe = vbeBuilder.Build().Object;
            var component1 = mockComponents[0].Object;
            var component2 = mockComponents[1].Object;
            var otherComponent = otherMockComponents[0].Object;

            var repository = new ProjectsRepository(vbe);
            var components = repository.Components().ToList();

            Assert.AreEqual(3, components.Count);
            Assert.Contains((component1.QualifiedModuleName, component1), components);
            Assert.Contains((component2.QualifiedModuleName, component2), components);
            Assert.Contains((otherComponent.QualifiedModuleName, otherComponent), components);
        }

        [Test()]
        public void ComponentsGetDisposedOnDisposal()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            projectBuilder.AddComponent("component2", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.BuildWIthAssignedProjectId();
            var mockComponents = projectBuilder.MockComponents;
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(otherProjectMock);
            var otherMockComponents = otherProjectBuilder.MockComponents;

            var vbe = vbeBuilder.Build().Object;

            var repository = new ProjectsRepository(vbe);
            repository.Dispose();

            foreach (var mock in mockComponents.Concat(otherMockComponents))
            {
                mock.Verify(m => m.Dispose(), Times.Once);
            }
        }

        [Test()]
        public void ComponentsReturnsEmptyCollectionAfterDisposal()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            projectBuilder.AddComponent("component2", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;

            var repository = new ProjectsRepository(vbe);
            repository.Dispose();
            var components = repository.Components().ToList();

            Assert.IsEmpty(components);
        }

        [Test()]
        public void ComponentsDoesNotReturnComponentsAddedToVbeBeforeRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;

            var repository = new ProjectsRepository(vbe);
            var newComponent = project.VBComponents.Add(ComponentType.ClassModule);
            var components = repository.Components().ToList();

            Assert.IsFalse(components.Contains((newComponent.QualifiedModuleName, newComponent)));
        }

        [Test()]
        public void ComponentsDoesNotReturnComponentsAddedToVbeAfterRefreshForOtherProjectId()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var otherProject = otherProjectMock.Object;

            var repository = new ProjectsRepository(vbe);
            var newComponent = project.VBComponents.Add(ComponentType.ClassModule);
            repository.Refresh(otherProject.ProjectId);
            var components = repository.Components().ToList();

            Assert.IsFalse(components.Contains((newComponent.QualifiedModuleName, newComponent)));
        }

        [Test()]
        public void ComponentsReturnsComponentsAddedToVbeAfterRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;

            var repository = new ProjectsRepository(vbe);
            var newComponent = project.VBComponents.Add(ComponentType.ClassModule);
            repository.Refresh();
            var components = repository.Components().ToList();

            Assert.Contains((newComponent.QualifiedModuleName, newComponent), components);
        }

        [Test()]
        public void ComponentsReturnsComponentsAddedToVbeAfterRefreshWIthPojectIdOfContainingProject()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;

            var repository = new ProjectsRepository(vbe);
            var newComponent = project.VBComponents.Add(ComponentType.ClassModule);
            repository.Refresh(project.ProjectId);
            var components = repository.Components().ToList();

            Assert.Contains((newComponent.QualifiedModuleName, newComponent), components);
        }

        [Test()]
        public void ComponentsReturnsRemovedComponentsBeforeRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            projectBuilder.AddComponent("component2", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.BuildWIthAssignedProjectId();
            var mockComponents = projectBuilder.MockComponents;
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var component2 = mockComponents[1].Object;


            var repository = new ProjectsRepository(vbe);
            project.VBComponents.Remove(component2);
            var components = repository.Components().ToList();

            Assert.Contains((component2.QualifiedModuleName, component2), components);
        }

        [Test()]
        public void ComponentsReturnsRemovedComponentsAfterRefreshForOtherProjectId()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            projectBuilder.AddComponent("component2", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.BuildWIthAssignedProjectId();
            var mockComponents = projectBuilder.MockComponents;
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var otherProject = otherProjectMock.Object;
            var component2 = mockComponents[1].Object;


            var repository = new ProjectsRepository(vbe);
            project.VBComponents.Remove(component2);
            repository.Refresh(otherProject.ProjectId);
            var components = repository.Components().ToList();

            Assert.Contains((component2.QualifiedModuleName, component2), components);
        }

        [Test()]
        public void ComponentsDoesNotReturnRemovedComponentsAfterRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            projectBuilder.AddComponent("component2", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.BuildWIthAssignedProjectId();
            var mockComponents = projectBuilder.MockComponents;
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var component2 = mockComponents[1].Object;

            var repository = new ProjectsRepository(vbe);
            project.VBComponents.Remove(component2);
            repository.Refresh();
            var components = repository.Components().ToList();

            Assert.IsFalse(components.Contains((component2.QualifiedModuleName, component2)));
        }

        [Test()]
        public void ComponentsDoesNotReturnRemovedComponentsAfterRefreshForProjectIdOfContainingProject()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            projectBuilder.AddComponent("component2", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.BuildWIthAssignedProjectId();
            var mockComponents = projectBuilder.MockComponents;
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var component2 = mockComponents[1].Object;

            var repository = new ProjectsRepository(vbe);
            project.VBComponents.Remove(component2);
            repository.Refresh(project.ProjectId);
            var components = repository.Components().ToList();

            Assert.IsFalse(components.Contains((component2.QualifiedModuleName, component2)));
        }

        [Test()]
        public void ComponentsGetDisposedOnRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            projectBuilder.AddComponent("component2", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.BuildWIthAssignedProjectId();
            var mockComponents = projectBuilder.MockComponents;
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(otherProjectMock);
            var otherMockComponents = otherProjectBuilder.MockComponents;

            var vbe = vbeBuilder.Build().Object;

            var repository = new ProjectsRepository(vbe);
            repository.Refresh();

            foreach (var mock in mockComponents.Concat(otherMockComponents))
            {
                mock.Verify(m => m.Dispose(), Times.Once);
            }
        }

        [Test()]
        public void ComponentsInProjectWithMatchingProjectIdGetDisposedOnRefreshForProjectId()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            projectBuilder.AddComponent("component2", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.BuildWIthAssignedProjectId();
            var mockComponents = projectBuilder.MockComponents;
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;

            var repository = new ProjectsRepository(vbe);
            repository.Refresh(project.ProjectId);

            foreach (var mock in mockComponents)
            {
                mock.Verify(m => m.Dispose(), Times.Once);
            }
        }

        [Test()]
        public void ComponentsInProjectWithOtherProjectIdDoNotGetDisposedOnRefreshForProjectId()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            projectBuilder.AddComponent("component2", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(otherProjectMock);
            var otherMockComponents = otherProjectBuilder.MockComponents;

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;

            var repository = new ProjectsRepository(vbe);
            repository.Refresh(project.ProjectId);

            foreach (var mock in otherMockComponents)
            {
                mock.Verify(m => m.Dispose(), Times.Never);
            }
        }

        [Test()]
        public void RemovedComponentsGetDisposedOnRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            projectBuilder.AddComponent("component2", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.BuildWIthAssignedProjectId();
            var mockComponents = projectBuilder.MockComponents;
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var component2Mock = mockComponents[1];
            var component2 = component2Mock.Object;

            var repository = new ProjectsRepository(vbe);
            project.VBComponents.Remove(component2);
            repository.Refresh();

            component2Mock.Verify(m => m.Dispose(), Times.Once);
        }

        [Test()]
        public void RemovedComponentsGetDisposedOnRefreshFroProjectIdOfFormerlyContainingProject()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            projectBuilder.AddComponent("component2", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.BuildWIthAssignedProjectId();
            var mockComponents = projectBuilder.MockComponents;
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var component2Mock = mockComponents[1];
            var component2 = component2Mock.Object;

            var repository = new ProjectsRepository(vbe);
            project.VBComponents.Remove(component2);
            repository.Refresh(project.ProjectId);

            component2Mock.Verify(m => m.Dispose(), Times.Once);
        }

        [Test()]
        public void RemovedComponentsDoNotGetDisposedOnRefreshForOtherProjectId()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            projectBuilder.AddComponent("component2", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.BuildWIthAssignedProjectId();
            var mockComponents = projectBuilder.MockComponents;
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.BuildWIthAssignedProjectId();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var otherProject = otherProjectMock.Object;
            var component2Mock = mockComponents[1];
            var component2 = component2Mock.Object;

            var repository = new ProjectsRepository(vbe);
            project.VBComponents.Remove(component2);
            repository.Refresh(otherProject.ProjectId);

            component2Mock.Verify(m => m.Dispose(), Times.Never);
        }

        [Test()]
        public void ComponentsForProjectIdTest()
        {
            Assert.Fail();
        }

        [Test()]
        public void ComponentTest()
        {
            Assert.Fail();
        }

        [Test()]
        public void CodeModulesTest()
        {
            Assert.Fail();
        }

        [Test()]
        public void CodeModuleTest()
        {
            Assert.Fail();
        }
    }
}