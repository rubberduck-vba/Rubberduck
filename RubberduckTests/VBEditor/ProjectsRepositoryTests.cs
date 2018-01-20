using System;
using System.Linq;
using System.Windows.Controls.Primitives;
using Moq;
using NUnit.Framework;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.VBEditor
{
    [TestFixture()]
    public class ProjectsRepositoryTests
    {
        private IProjectsRepository TestRepository(IVBE vbe)
        {
            return new ProjectsRepository(vbe);
        }

        [Test()]
        public void ProjectsCollectionReturnsTheOneFromTheVbePassedIn()
        {
            var vbe = new MockVbeBuilder().Build().Object;
            var repository = TestRepository(vbe);

            var vbePojects = vbe.VBProjects;
            var repositoryProjects = repository.ProjectsCollection();

            Assert.AreEqual(vbePojects, repositoryProjects);
        }

        [Test()]
        public void ProjectsCollectionGetsDisposedOnDisposal()
        {
            var vbeBuilder = new MockVbeBuilder();
            var vbe = vbeBuilder.Build().Object;
            var repository = TestRepository(vbe);

            repository.Dispose();

            var projectsMock = vbeBuilder.MockProjectsCollection;
            projectsMock.Verify(m => m.Dispose(), Times.Once);
        }

        [Test()]
        public void ProjectsCollectionReturnsNullAfterDisposal()
        {
            var vbe = new MockVbeBuilder().Build().Object;
            var repository = TestRepository(vbe);

            repository.Dispose();

            Assert.IsNull(repository.ProjectsCollection());
        }

        [Test()]
        public void ProjectsReturnsProjectsOnVbe()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            var projectMock = projectBuilder.Build();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var otherProject = otherProjectMock.Object;

            var repository = TestRepository(vbe);
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
            var projectMock = projectBuilder.Build();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;

            var repository = TestRepository(vbe);
            repository.Dispose();

            projectMock.Verify(m => m.Dispose(), Times.Once);
            otherProjectMock.Verify(m => m.Dispose(), Times.Once);
        }

        [Test()]
        public void ProjectsReturnsEmptyCollectionAfterDisposal()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            var projectMock = projectBuilder.Build();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;

            var repository = TestRepository(vbe);
            repository.Dispose();
            var projects = repository.Projects().ToList();

            Assert.IsEmpty(projects);
        }

        [Test()]
        public void ProjectsDoesNotReturnProjectsAddedToVbeWithoutRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            var projectMock = projectBuilder.Build();
            vbeBuilder.AddProject(projectMock);

            var vbe = vbeBuilder.Build().Object;

            var repository = TestRepository(vbe);
            var otherProject = vbe.VBProjects.Add(ProjectType.HostProject);
            var projects = repository.Projects().ToList();

            Assert.IsFalse(projects.Contains((otherProject.ProjectId, otherProject)));
        }

        [Test()]
        public void ProjectsReturnsProjectsAddedToVbeAfterRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            var projectMock = projectBuilder.Build();
            vbeBuilder.AddProject(projectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;

            var repository = TestRepository(vbe);
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
            var projectMock = projectBuilder.Build();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var otherProject = otherProjectMock.Object;

            var repository = TestRepository(vbe);
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
            var projectMock = projectBuilder.Build();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var otherProject = otherProjectMock.Object;

            var repository = TestRepository(vbe);
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
            var projectMock = projectBuilder.Build();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;

            var repository = TestRepository(vbe);
            repository.Refresh();

            projectMock.Verify(m => m.Dispose(), Times.Once);
            otherProjectMock.Verify(m => m.Dispose(), Times.Once);
        }

        [Test()]
        public void RemovedProjectsGetDisposedOnRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            var projectMock = projectBuilder.Build();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var otherProject = otherProjectMock.Object;

            var repository = TestRepository(vbe);
            vbe.VBProjects.Remove(otherProject);
            repository.Refresh();

            otherProjectMock.Verify(m => m.Dispose(), Times.Once);
        }

        [Test()]
        public void ProjectReturnsProjectWithMatchingProjectId()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            var projectMock = projectBuilder.Build();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var otherProject = otherProjectMock.Object;

            var repository = TestRepository(vbe);
            var returnedProject = repository.Project(otherProject.ProjectId);

            Assert.AreEqual(otherProject, returnedProject);
        }

        [Test()]
        public void ProjectReturnsNullForUnknownProjectId()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            var projectMock = projectBuilder.Build();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;

            var repository = TestRepository(vbe);
            var returnedProject = repository.Project(new Guid().ToString());

            Assert.IsNull(returnedProject);
        }

        [Test()]
        public void ProjectReturnsNullAfterDisposal()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            var projectMock = projectBuilder.Build();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var otherProject = otherProjectMock.Object;

            var repository = TestRepository(vbe);
            repository.Dispose();
            var returnedProject = repository.Project(otherProject.ProjectId);

            Assert.IsNull(returnedProject);
        }

        [Test()]
        public void ProjectReturnsNullForProjectIdOfAddedProjectBeforeRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            var projectMock = projectBuilder.Build();
            vbeBuilder.AddProject(projectMock);

            var vbe = vbeBuilder.Build().Object;

            var repository = TestRepository(vbe);
            var otherProject = vbe.VBProjects.Add(ProjectType.HostProject);
            var returnedProject = repository.Project(otherProject.ProjectId);

            Assert.IsNull(returnedProject);
        }

        [Test()]
        public void ProjectReturnsAddedProjectWithMatchingProjectIdAfterRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            var projectMock = projectBuilder.Build();
            vbeBuilder.AddProject(projectMock);

            var vbe = vbeBuilder.Build().Object;

            var repository = TestRepository(vbe);
            var otherProject = vbe.VBProjects.Add(ProjectType.HostProject);
            repository.Refresh();
            var returnedProject = repository.Project(otherProject.ProjectId);

            Assert.AreEqual(otherProject, returnedProject);
        }

        [Test()]
        public void ProjectReturnsRemovedProjectWithMatchingProjectIdBeforeRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            var projectMock = projectBuilder.Build();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var otherProject = otherProjectMock.Object;

            var repository = TestRepository(vbe);
            vbe.VBProjects.Remove(otherProject);
            var returnedProject = repository.Project(otherProject.ProjectId);

            Assert.AreEqual(otherProject, returnedProject);
        }

        [Test()]
        public void ProjectReturnsNullForProjectIdOfRemovedProjectAfterRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            var projectMock = projectBuilder.Build();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var otherProject = otherProjectMock.Object;

            var repository = TestRepository(vbe);
            vbe.VBProjects.Remove(otherProject);
            repository.Refresh();
            var returnedProject = repository.Project(otherProject.ProjectId);

            Assert.IsNull(returnedProject);
        }

        [Test()]
        public void ComponentsCollectionsOfProjectsGetDisposedOnDisposal()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            var projectMock = projectBuilder.Build();
            vbeBuilder.AddProject(projectMock);
            var componentsCollectionMock = projectBuilder.MockVBComponents;
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);
            var otherComponentsCollectionMock = otherProjectBuilder.MockVBComponents;

            var vbe = vbeBuilder.Build().Object;

            var repository = TestRepository(vbe);
            repository.Dispose();

            componentsCollectionMock.Verify(m => m.Dispose(), Times.Once);
            otherComponentsCollectionMock.Verify(m => m.Dispose(), Times.Once);
        }

        [Test()]
        public void ComponentsCollectionsOfProjectsGetDisposedOnRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            var projectMock = projectBuilder.Build();
            vbeBuilder.AddProject(projectMock);
            var componentsCollectionMock = projectBuilder.MockVBComponents;
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);
            var otherComponentsCollectionMock = otherProjectBuilder.MockVBComponents;

            var vbe = vbeBuilder.Build().Object;

            var repository = TestRepository(vbe);
            repository.Refresh();

            componentsCollectionMock.Verify(m => m.Dispose(), Times.Once);
            otherComponentsCollectionMock.Verify(m => m.Dispose(), Times.Once);
        }

        [Test()]
        public void ComponentsCollectionsOfRemovedProjectsGetDisposedOnRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            var projectMock = projectBuilder.Build();
            vbeBuilder.AddProject(projectMock);
            var componentsCollectionMock = projectBuilder.MockVBComponents;
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var otherProject = otherProjectMock.Object;

            var repository = TestRepository(vbe);
            vbe.VBProjects.Remove(otherProject);
            repository.Refresh();

            componentsCollectionMock.Verify(m => m.Dispose(), Times.Once);
        }

        [Test()]
        public void ComponentsCollectionReturnsComponentsCollectionOfProjectWithMatchingProjectId()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            var projectMock = projectBuilder.Build();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;

            var repository = TestRepository(vbe);
            var returnedCollection = repository.ComponentsCollection(project.ProjectId);

            Assert.AreEqual(project.VBComponents, returnedCollection);
        }

        [Test()]
        public void ComponentsCollectionReturnsNullForUnknownProjectId()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            var projectMock = projectBuilder.Build();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;

            var repository = TestRepository(vbe);
            var returnedCollection = repository.ComponentsCollection(new Guid().ToString());

            Assert.IsNull(returnedCollection);
        }

        [Test()]
        public void ComponentsCollectionReturnsNullAfterDisposal()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            var projectMock = projectBuilder.Build();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;

            var repository = TestRepository(vbe);
            repository.Dispose();
            var returnedCollection = repository.ComponentsCollection(project.ProjectId);

            Assert.IsNull(returnedCollection);
        }

        [Test()]
        public void ComponentsCollectionReturnsNullForProjectIdOfAddedProjectBeforeRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            var projectMock = projectBuilder.Build();
            vbeBuilder.AddProject(projectMock);

            var vbe = vbeBuilder.Build().Object;

            var repository = TestRepository(vbe);
            var otherProject = vbe.VBProjects.Add(ProjectType.HostProject);
            var returnedCollection = repository.ComponentsCollection(otherProject.ProjectId);

            Assert.IsNull(returnedCollection);
        }

        [Test()]
        public void ComponentCollectionReturnsComponentsCollectionOfAddedProjectWithMatchingProjectIdAfterRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            var projectMock = projectBuilder.Build();
            vbeBuilder.AddProject(projectMock);

            var vbe = vbeBuilder.Build().Object;

            var repository = TestRepository(vbe);
            var otherProject = vbe.VBProjects.Add(ProjectType.HostProject);
            repository.Refresh();
            var returnedCollection = repository.ComponentsCollection(otherProject.ProjectId);

            Assert.AreEqual(otherProject.VBComponents, returnedCollection);
        }

        [Test()]
        public void ComponentsCollectionReturnsComponentsCollectionOfRemovedProjectWithMatchingProjectIdBeforeRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            var projectMock = projectBuilder.Build();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var otherProject = otherProjectMock.Object;

            var repository = TestRepository(vbe);
            vbe.VBProjects.Remove(otherProject);
            var returnedCollection = repository.ComponentsCollection(otherProject.ProjectId);

            Assert.AreEqual(otherProject.VBComponents, returnedCollection);
        }

        [Test()]
        public void ComponentsCollectionReturnsNullForProjectIdOfRemovedProjectAfterRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            var projectMock = projectBuilder.Build();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var otherProject = otherProjectMock.Object;

            var repository = TestRepository(vbe);
            vbe.VBProjects.Remove(otherProject);
            repository.Refresh();
            var returnedCollection = repository.ComponentsCollection(otherProject.ProjectId);

            Assert.IsNull(returnedCollection);
        }

        [Test()]
        public void ComponentsCollectionOfProjectWithMatchingProjectIdGetsDisposedOnRefreshForProjectId()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            var projectMock = projectBuilder.Build();
            vbeBuilder.AddProject(projectMock);
            var componentsCollectionMock = projectBuilder.MockVBComponents;
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;

            var repository = TestRepository(vbe);
            repository.Refresh(project.ProjectId);

            componentsCollectionMock.Verify(m => m.Dispose(), Times.Once);
        }

        [Test()]
        public void ComponentsCollectionsOfProjectsWithNonMatchingProjectIdDoNotGetDisposedOnRefreshForProjectId()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            var projectMock = projectBuilder.Build();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);
            var otherComponentsCollectionMock = otherProjectBuilder.MockVBComponents;

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;

            var repository = TestRepository(vbe);
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
            var projectMock = projectBuilder.Build();
            var mockComponents = projectBuilder.MockComponents;
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);
            var otherMockComponents = otherProjectBuilder.MockComponents;

            var vbe = vbeBuilder.Build().Object;
            var component1 = mockComponents[0].Object;
            var component2 = mockComponents[1].Object;
            var otherComponent = otherMockComponents[0].Object;

            var repository = TestRepository(vbe);
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
            var projectMock = projectBuilder.Build();
            var mockComponents = projectBuilder.MockComponents;
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);
            var otherMockComponents = otherProjectBuilder.MockComponents;

            var vbe = vbeBuilder.Build().Object;

            var repository = TestRepository(vbe);
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
            var projectMock = projectBuilder.Build();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;

            var repository = TestRepository(vbe);
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
            var projectMock = projectBuilder.Build();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;

            var repository = TestRepository(vbe);
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
            var projectMock = projectBuilder.Build();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var otherProject = otherProjectMock.Object;

            var repository = TestRepository(vbe);
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
            var projectMock = projectBuilder.Build();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;

            var repository = TestRepository(vbe);
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
            var projectMock = projectBuilder.Build();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;

            var repository = TestRepository(vbe);
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
            var projectMock = projectBuilder.Build();
            var mockComponents = projectBuilder.MockComponents;
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var component2 = mockComponents[1].Object;


            var repository = TestRepository(vbe);
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
            var projectMock = projectBuilder.Build();
            var mockComponents = projectBuilder.MockComponents;
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var otherProject = otherProjectMock.Object;
            var component2 = mockComponents[1].Object;


            var repository = TestRepository(vbe);
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
            var projectMock = projectBuilder.Build();
            var mockComponents = projectBuilder.MockComponents;
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var component2 = mockComponents[1].Object;

            var repository = TestRepository(vbe);
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
            var projectMock = projectBuilder.Build();
            var mockComponents = projectBuilder.MockComponents;
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var component2 = mockComponents[1].Object;

            var repository = TestRepository(vbe);
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
            var projectMock = projectBuilder.Build();
            var mockComponents = projectBuilder.MockComponents;
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);
            var otherMockComponents = otherProjectBuilder.MockComponents;

            var vbe = vbeBuilder.Build().Object;

            var repository = TestRepository(vbe);
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
            var projectMock = projectBuilder.Build();
            var mockComponents = projectBuilder.MockComponents;
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;

            var repository = TestRepository(vbe);
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
            var projectMock = projectBuilder.Build();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);
            var otherMockComponents = otherProjectBuilder.MockComponents;

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;

            var repository = TestRepository(vbe);
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
            var projectMock = projectBuilder.Build();
            var mockComponents = projectBuilder.MockComponents;
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var component2Mock = mockComponents[1];
            var component2 = component2Mock.Object;

            var repository = TestRepository(vbe);
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
            var projectMock = projectBuilder.Build();
            var mockComponents = projectBuilder.MockComponents;
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var component2Mock = mockComponents[1];
            var component2 = component2Mock.Object;

            var repository = TestRepository(vbe);
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
            var projectMock = projectBuilder.Build();
            var mockComponents = projectBuilder.MockComponents;
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var otherProject = otherProjectMock.Object;
            var component2Mock = mockComponents[1];
            var component2 = component2Mock.Object;

            var repository = TestRepository(vbe);
            project.VBComponents.Remove(component2);
            repository.Refresh(otherProject.ProjectId);

            component2Mock.Verify(m => m.Dispose(), Times.Never);
        }

        [Test()]
        public void ComponentsForProjectIdReturnsComponentsOnProjectWithMatchingProjectId()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            projectBuilder.AddComponent("component2", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.Build();
            var mockComponents = projectBuilder.MockComponents;
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var component1 = mockComponents[0].Object;
            var component2 = mockComponents[1].Object;

            var repository = TestRepository(vbe);
            var components = repository.Components(project.ProjectId).ToList();

            Assert.AreEqual(2, components.Count);
            Assert.Contains((component1.QualifiedModuleName, component1), components);
            Assert.Contains((component2.QualifiedModuleName, component2), components);
        }

        [Test()]
        public void ComponentsForProjectIdReturnsEmptyCollectionAfterDisposal()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            projectBuilder.AddComponent("component2", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.Build();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;

            var repository = TestRepository(vbe);
            repository.Dispose();
            var components = repository.Components(project.ProjectId).ToList();

            Assert.IsEmpty(components);
        }

        [Test()]
        public void ComponentsFroProjectIdDoesNotReturnComponentsAddedToVbeBeforeRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.Build();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;

            var repository = TestRepository(vbe);
            var newComponent = project.VBComponents.Add(ComponentType.ClassModule);
            var components = repository.Components(project.ProjectId).ToList();

            Assert.IsFalse(components.Contains((newComponent.QualifiedModuleName, newComponent)));
        }

        [Test()]
        public void ComponentsForProjectIdDoesNotReturnComponentsAddedToVbeAfterRefreshForOtherProjectId()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.Build();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var otherProject = otherProjectMock.Object;

            var repository = TestRepository(vbe);
            var newComponent = project.VBComponents.Add(ComponentType.ClassModule);
            repository.Refresh(otherProject.ProjectId);
            var components = repository.Components(project.ProjectId).ToList();

            Assert.IsFalse(components.Contains((newComponent.QualifiedModuleName, newComponent)));
        }

        [Test()]
        public void ComponentsForProjectIdReturnsComponentsAddedToVbeAfterRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.Build();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;

            var repository = TestRepository(vbe);
            var newComponent = project.VBComponents.Add(ComponentType.ClassModule);
            repository.Refresh();
            var components = repository.Components(project.ProjectId).ToList();

            Assert.Contains((newComponent.QualifiedModuleName, newComponent), components);
        }

        [Test()]
        public void ComponentsForProjectIdReturnsComponentsAddedToVbeAfterRefreshWIthPojectIdOfContainingProject()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.Build();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;

            var repository = TestRepository(vbe);
            var newComponent = project.VBComponents.Add(ComponentType.ClassModule);
            repository.Refresh(project.ProjectId);
            var components = repository.Components(project.ProjectId).ToList();

            Assert.Contains((newComponent.QualifiedModuleName, newComponent), components);
        }

        [Test()]
        public void ComponentsForProjectIdReturnsRemovedComponentsBeforeRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            projectBuilder.AddComponent("component2", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.Build();
            var mockComponents = projectBuilder.MockComponents;
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var component2 = mockComponents[1].Object;


            var repository = TestRepository(vbe);
            project.VBComponents.Remove(component2);
            var components = repository.Components(project.ProjectId).ToList();

            Assert.Contains((component2.QualifiedModuleName, component2), components);
        }

        [Test()]
        public void ComponentsForProjectIdReturnsRemovedComponentsAfterRefreshForOtherProjectId()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            projectBuilder.AddComponent("component2", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.Build();
            var mockComponents = projectBuilder.MockComponents;
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var otherProject = otherProjectMock.Object;
            var component2 = mockComponents[1].Object;


            var repository = TestRepository(vbe);
            project.VBComponents.Remove(component2);
            repository.Refresh(otherProject.ProjectId);
            var components = repository.Components(project.ProjectId).ToList();

            Assert.Contains((component2.QualifiedModuleName, component2), components);
        }

        [Test()]
        public void ComponentsForProjectIdDoesNotReturnRemovedComponentsAfterRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            projectBuilder.AddComponent("component2", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.Build();
            var mockComponents = projectBuilder.MockComponents;
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var component2 = mockComponents[1].Object;

            var repository = TestRepository(vbe);
            project.VBComponents.Remove(component2);
            repository.Refresh();
            var components = repository.Components(project.ProjectId).ToList();

            Assert.IsFalse(components.Contains((component2.QualifiedModuleName, component2)));
        }

        [Test()]
        public void ComponentsForProjectIdDoesNotReturnRemovedComponentsAfterRefreshForProjectIdOfContainingProject()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            projectBuilder.AddComponent("component2", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.Build();
            var mockComponents = projectBuilder.MockComponents;
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var component2 = mockComponents[1].Object;

            var repository = TestRepository(vbe);
            project.VBComponents.Remove(component2);
            repository.Refresh(project.ProjectId);
            var components = repository.Components(project.ProjectId).ToList();

            Assert.IsFalse(components.Contains((component2.QualifiedModuleName, component2)));
        }

        [Test()]
        public void ComponentReturnsComponentWithMatchingQmn()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            projectBuilder.AddComponent("component2", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.Build();
            var mockComponents = projectBuilder.MockComponents;
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var component2 = mockComponents[1].Object;

            var repository = TestRepository(vbe);
            var returnedComponent = repository.Component(component2.QualifiedModuleName);

            Assert.AreEqual(component2, returnedComponent);
        }

        [Test()]
        public void ComponentReturnsNullForUnknownQmn()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            projectBuilder.AddComponent("component2", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.Build();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;

            var repository = TestRepository(vbe);
            var newQmn = new QualifiedModuleName(String.Empty, String.Empty, "newComponent");
            var returnedComponent = repository.Component(newQmn);

            Assert.IsNull(returnedComponent);
        }

        [Test()]
        public void ComponentReturnsNullAfterDisposal()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            projectBuilder.AddComponent("component2", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.Build();
            var mockComponents = projectBuilder.MockComponents;
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var component2 = mockComponents[1].Object;

            var repository = TestRepository(vbe);
            repository.Dispose();
            var returnedComponent = repository.Component(component2.QualifiedModuleName);

            Assert.IsNull(returnedComponent);
        }

        [Test()]
        public void ComponentReturnsNullForQmnOfAddedComponentBeforeRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.Build();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;

            var repository = TestRepository(vbe);
            var newComponent = project.VBComponents.Add(ComponentType.ClassModule);
            var returnedComponent = repository.Component(newComponent.QualifiedModuleName);

            Assert.IsNull(returnedComponent);
        }

        [Test()]
        public void ComponentReturnsNullForQmnOfAddedComponentAfterRefreshForOtherProjectId()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.Build();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var otherProject = otherProjectMock.Object;

            var repository = TestRepository(vbe);
            var newComponent = project.VBComponents.Add(ComponentType.ClassModule);
            repository.Refresh(otherProject.ProjectId);
            var returnedComponent = repository.Component(newComponent.QualifiedModuleName);

            Assert.IsNull(returnedComponent);
        }

        [Test()]
        public void ComponentReturnsAddedComponentWithMatchingQmnAfterRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.Build();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;

            var repository = TestRepository(vbe);
            var newComponent = project.VBComponents.Add(ComponentType.ClassModule);
            repository.Refresh();
            var returnedComponent = repository.Component(newComponent.QualifiedModuleName);

            Assert.AreEqual(newComponent, returnedComponent);
        }

        [Test()]
        public void ComponentReturnsAddedComponentWithMatchingQmnAfterRefreshForProjectIdOfContaiingProject()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.Build();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;

            var repository = TestRepository(vbe);
            var newComponent = project.VBComponents.Add(ComponentType.ClassModule);
            repository.Refresh(project.ProjectId);
            var returnedComponent = repository.Component(newComponent.QualifiedModuleName);

            Assert.AreEqual(newComponent, returnedComponent);
        }

        [Test()]
        public void ComponentReturnsRemovedComponentWithMatchingQmnBeforeRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            projectBuilder.AddComponent("component2", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.Build();
            var mockComponents = projectBuilder.MockComponents;
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var component2 = mockComponents[1].Object;

            var repository = TestRepository(vbe);
            project.VBComponents.Remove(component2);
            var returnedComponent = repository.Component(component2.QualifiedModuleName);

            Assert.AreEqual(component2, returnedComponent);
        }

        [Test()]
        public void ComponentReturnsRemovedComponentWithMatchingQmnAfterRefreshFroOtherProjectId()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            projectBuilder.AddComponent("component2", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.Build();
            var mockComponents = projectBuilder.MockComponents;
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var otherProject = otherProjectMock.Object;
            var component2 = mockComponents[1].Object;

            var repository = TestRepository(vbe);
            project.VBComponents.Remove(component2);
            repository.Refresh(otherProject.ProjectId);
            var returnedComponent = repository.Component(component2.QualifiedModuleName);

            Assert.AreEqual(component2, returnedComponent);
        }

        [Test()]
        public void ComponentReturnsNullForQmnOfRemovedComponentAfterRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            projectBuilder.AddComponent("component2", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.Build();
            var mockComponents = projectBuilder.MockComponents;
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var component2 = mockComponents[1].Object;

            var repository = TestRepository(vbe);
            project.VBComponents.Remove(component2);
            repository.Refresh();
            var returnedComponent = repository.Component(component2.QualifiedModuleName);

            Assert.IsNull(returnedComponent);
        }

        [Test()]
        public void ComponentReturnsNullForQmnOfRemovedComponentAfterRefreshForProjectIdOfContainingProject()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            projectBuilder.AddComponent("component2", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.Build();
            var mockComponents = projectBuilder.MockComponents;
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var component2 = mockComponents[1].Object;

            var repository = TestRepository(vbe);
            project.VBComponents.Remove(component2);
            repository.Refresh(project.ProjectId);
            var returnedComponent = repository.Component(component2.QualifiedModuleName);

            Assert.IsNull(returnedComponent);
        }

        [Test()]
        public void CodeModulesReturnsCodeModulesOnVbeWithQmns()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            projectBuilder.AddComponent("component2", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.Build();
            var mockComponents = projectBuilder.MockComponents;
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);
            var otherMockComponents = otherProjectBuilder.MockComponents;

            var vbe = vbeBuilder.Build().Object;
            var component1 = mockComponents[0].Object;
            var component2 = mockComponents[1].Object;
            var otherComponent = otherMockComponents[0].Object;

            var repository = TestRepository(vbe);
            var codeModules = repository.CodeModules().ToList();

            Assert.AreEqual(3, codeModules.Count);
            Assert.Contains((component1.QualifiedModuleName, component1.CodeModule), codeModules);
            Assert.Contains((component2.QualifiedModuleName, component2.CodeModule), codeModules);
            Assert.Contains((otherComponent.QualifiedModuleName, otherComponent.CodeModule), codeModules);
        }

        [Test()]
        public void CodeModulesGetDisposedOnDisposal()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            projectBuilder.AddComponent("component2", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.Build();
            var mockCodeModules = projectBuilder.MockCodeModules;
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);
            var otherMockCodeModules = otherProjectBuilder.MockCodeModules;

            var vbe = vbeBuilder.Build().Object;

            var repository = TestRepository(vbe);
            repository.Dispose();

            foreach (var mock in mockCodeModules.Concat(otherMockCodeModules))
            {
                mock.Verify(m => m.Dispose(), Times.Once);
            }
        }

        [Test()]
        public void CodeModulesReturnsEmptyCollectionAfterDisposal()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            projectBuilder.AddComponent("component2", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.Build();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;

            var repository = TestRepository(vbe);
            repository.Dispose();
            var codeModules = repository.CodeModules().ToList();

            Assert.IsEmpty(codeModules);
        }

        [Test()]
        public void CodeModulesDoesNotReturnCodeModulesAddedToVbeBeforeRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.Build();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;

            var repository = TestRepository(vbe);
            var newComponent = project.VBComponents.Add(ComponentType.ClassModule);
            var codeModules = repository.CodeModules().ToList();

            Assert.IsFalse(codeModules.Contains((newComponent.QualifiedModuleName, newComponent.CodeModule)));
        }

        [Test()]
        public void CodeModulesDoesNotReturnCodeModulesAddedToVbeAfterRefreshForOtherProjectId()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.Build();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var otherProject = otherProjectMock.Object;

            var repository = TestRepository(vbe);
            var newComponent = project.VBComponents.Add(ComponentType.ClassModule);
            repository.Refresh(otherProject.ProjectId);
            var codeModules = repository.CodeModules().ToList();

            Assert.IsFalse(codeModules.Contains((newComponent.QualifiedModuleName, newComponent.CodeModule)));
        }

        [Test()]
        public void CodeModulesReturnsCodeModulesAddedToVbeAfterRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.Build();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;

            var repository = TestRepository(vbe);
            var newComponent = project.VBComponents.Add(ComponentType.ClassModule);
            repository.Refresh();
            var codeModules = repository.CodeModules().ToList();

            Assert.Contains((newComponent.QualifiedModuleName, newComponent.CodeModule), codeModules);
        }

        [Test()]
        public void CodeModulesReturnsCodeModulesAddedToVbeAfterRefreshWIthPojectIdOfContainingProject()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.Build();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;

            var repository = TestRepository(vbe);
            var newComponent = project.VBComponents.Add(ComponentType.ClassModule);
            repository.Refresh(project.ProjectId);
            var codeModules = repository.CodeModules().ToList();

            Assert.Contains((newComponent.QualifiedModuleName, newComponent.CodeModule), codeModules);
        }

        [Test()]
        public void CodeModulesReturnsRemovedCodeModulesBeforeRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            projectBuilder.AddComponent("component2", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.Build();
            var mockComponents = projectBuilder.MockComponents;
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var component2 = mockComponents[1].Object;


            var repository = TestRepository(vbe);
            project.VBComponents.Remove(component2);
            var codeModules = repository.CodeModules().ToList();

            Assert.Contains((component2.QualifiedModuleName, component2.CodeModule), codeModules);
        }

        [Test()]
        public void CodeModulesReturnsRemovedCodeModulesAfterRefreshForOtherProjectId()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            projectBuilder.AddComponent("component2", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.Build();
            var mockComponents = projectBuilder.MockComponents;
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var otherProject = otherProjectMock.Object;
            var component2 = mockComponents[1].Object;

            var repository = TestRepository(vbe);
            project.VBComponents.Remove(component2);
            repository.Refresh(otherProject.ProjectId);
            var codeModules = repository.CodeModules().ToList();

            Assert.Contains((component2.QualifiedModuleName, component2.CodeModule), codeModules);
        }

        [Test()]
        public void CodeModulesDoesNotReturnRemovedCodeModulesAfterRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            projectBuilder.AddComponent("component2", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.Build();
            var mockComponents = projectBuilder.MockComponents;
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var component2 = mockComponents[1].Object;

            var repository = TestRepository(vbe);
            project.VBComponents.Remove(component2);
            repository.Refresh();
            var codeModules = repository.CodeModules().ToList();

            Assert.IsFalse(codeModules.Contains((component2.QualifiedModuleName, component2.CodeModule)));
        }

        [Test()]
        public void CodeModulesDoesNotReturnRemovedCodeModulesAfterRefreshForProjectIdOfContainingProject()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            projectBuilder.AddComponent("component2", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.Build();
            var mockComponents = projectBuilder.MockComponents;
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var component2 = mockComponents[1].Object;

            var repository = TestRepository(vbe);
            project.VBComponents.Remove(component2);
            repository.Refresh(project.ProjectId);
            var codeModules = repository.CodeModules().ToList();

            Assert.IsFalse(codeModules.Contains((component2.QualifiedModuleName, component2.CodeModule)));
        }

        [Test()]
        public void CodeModulesGetDisposedOnRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            projectBuilder.AddComponent("component2", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.Build();
            var mockCodeModules = projectBuilder.MockCodeModules;
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);
            var otherMockCodeModules = otherProjectBuilder.MockCodeModules;

            var vbe = vbeBuilder.Build().Object;

            var repository = TestRepository(vbe);
            repository.Refresh();

            foreach (var mock in mockCodeModules.Concat(otherMockCodeModules))
            {
                mock.Verify(m => m.Dispose(), Times.Once);    
            }
        }

        [Test()]
        public void CodeModulesInProjectWithMatchingProjectIdGetDisposedOnRefreshForProjectId()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            projectBuilder.AddComponent("component2", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.Build();
            var mockCodeModules = projectBuilder.MockCodeModules;
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;

            var repository = TestRepository(vbe);
            repository.Refresh(project.ProjectId);

            foreach (var mock in mockCodeModules)
            {
                mock.Verify(m => m.Dispose(), Times.Once);
            }
        }

        [Test()]
        public void CodeModulesInProjectWithOtherProjectIdDoNotGetDisposedOnRefreshForProjectId()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            projectBuilder.AddComponent("component2", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.Build();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);
            var otherMockCodeModules = otherProjectBuilder.MockCodeModules;

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;

            var repository = TestRepository(vbe);
            repository.Refresh(project.ProjectId);

            foreach (var mock in otherMockCodeModules)
            {
                mock.Verify(m => m.Dispose(), Times.Never);
            }
        }

        [Test()]
        public void RemovedCodeModulesGetDisposedOnRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            projectBuilder.AddComponent("component2", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.Build();
            var mockCodeModules = projectBuilder.MockCodeModules;
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var codeModule2Mock = mockCodeModules[1];
            var component2 = codeModule2Mock.Object.Parent;

            var repository = TestRepository(vbe);
            project.VBComponents.Remove(component2);
            repository.Refresh();

            codeModule2Mock.Verify(m => m.Dispose(), Times.Once);
        }

        [Test()]
        public void RemovedCodeModulesGetDisposedOnRefreshForProjectIdOfFormerlyContainingProject()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            projectBuilder.AddComponent("component2", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.Build();
            var mockCodeModules = projectBuilder.MockCodeModules;
            var mockComponents = projectBuilder.MockComponents;
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var codeModule2Mock = mockCodeModules[1];
            var component2 = mockComponents[1].Object;

            var repository = TestRepository(vbe);
            project.VBComponents.Remove(component2);
            repository.Refresh(project.ProjectId);

            codeModule2Mock.Verify(m => m.Dispose(), Times.Once);
        }

        [Test()]
        public void RemovedCodeModulesDoNotGetDisposedOnRefreshForOtherProjectId()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            projectBuilder.AddComponent("component2", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.Build();
            var mockCodeModules = projectBuilder.MockCodeModules;
            var mockComponents = projectBuilder.MockComponents;
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var otherProject = otherProjectMock.Object;
            var codeModule2Mock = mockCodeModules[1];
            var component2 = mockComponents[1].Object;

            var repository = TestRepository(vbe);
            project.VBComponents.Remove(component2);
            repository.Refresh(otherProject.ProjectId);

            codeModule2Mock.Verify(m => m.Dispose(), Times.Never);
        }

        [Test()]
        public void CodeModulesForProjectIdReturnsCodeModulesOnProjectWithMatchingProjectId()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            projectBuilder.AddComponent("component2", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.Build();
            var mockCodeModules = projectBuilder.MockCodeModules;
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var codeModule1 = mockCodeModules[0].Object;
            var codeModule2 = mockCodeModules[1].Object;

            var repository = TestRepository(vbe);
            var codeModules = repository.CodeModules(project.ProjectId).ToList();

            Assert.AreEqual(2, codeModules.Count);
            Assert.Contains((codeModule1.QualifiedModuleName, codeModule1), codeModules);
            Assert.Contains((codeModule2.QualifiedModuleName, codeModule2), codeModules);
        }

        [Test()]
        public void CodeModulesForProjectIdReturnsEmptyCollectionAfterDisposal()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            projectBuilder.AddComponent("component2", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.Build();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;

            var repository = TestRepository(vbe);
            repository.Dispose();
            var codeModules = repository.CodeModules(project.ProjectId).ToList();

            Assert.IsEmpty(codeModules);
        }

        [Test()]
        public void CodeModulesFroProjectIdDoesNotReturnCodeModulesAddedToVbeBeforeRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.Build();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;

            var repository = TestRepository(vbe);
            var newComponent = project.VBComponents.Add(ComponentType.ClassModule);
            var codeModules = repository.CodeModules(project.ProjectId).ToList();

            Assert.IsFalse(codeModules.Contains((newComponent.QualifiedModuleName, newComponent.CodeModule)));
        }

        [Test()]
        public void CodeModulesForProjectIdDoesNotReturnCodeModulesAddedToVbeAfterRefreshForOtherProjectId()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.Build();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var otherProject = otherProjectMock.Object;

            var repository = TestRepository(vbe);
            var newComponent = project.VBComponents.Add(ComponentType.ClassModule);
            repository.Refresh(otherProject.ProjectId);
            var codeModules = repository.CodeModules(project.ProjectId).ToList();

            Assert.IsFalse(codeModules.Contains((newComponent.QualifiedModuleName, newComponent.CodeModule)));
        }

        [Test()]
        public void CodeModulesForProjectIdReturnsCodeModulesAddedToVbeAfterRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.Build();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;

            var repository = TestRepository(vbe);
            var newComponent = project.VBComponents.Add(ComponentType.ClassModule);
            repository.Refresh();
            var codeModules = repository.CodeModules(project.ProjectId).ToList();

            Assert.Contains((newComponent.QualifiedModuleName, newComponent.CodeModule), codeModules);
        }

        [Test()]
        public void CodeModulesForProjectIdReturnsCodeModulesAddedToVbeAfterRefreshWIthPojectIdOfContainingProject()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.Build();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;

            var repository = TestRepository(vbe);
            var newComponent = project.VBComponents.Add(ComponentType.ClassModule);
            repository.Refresh(project.ProjectId);
            var codeModules = repository.CodeModules(project.ProjectId).ToList();

            Assert.Contains((newComponent.QualifiedModuleName, newComponent.CodeModule), codeModules);
        }

        [Test()]
        public void CodeModulesForProjectIdReturnsRemovedCodeModulesBeforeRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            projectBuilder.AddComponent("component2", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.Build();
            var mockComponents = projectBuilder.MockComponents;
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var component2 = mockComponents[1].Object;


            var repository = TestRepository(vbe);
            project.VBComponents.Remove(component2);
            var codeModules = repository.CodeModules(project.ProjectId).ToList();

            Assert.Contains((component2.QualifiedModuleName, component2.CodeModule), codeModules);
        }

        [Test()]
        public void CodeModulesForProjectIdReturnsRemovedCodeModulesAfterRefreshForOtherProjectId()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            projectBuilder.AddComponent("component2", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.Build();
            var mockCodeModules = projectBuilder.MockComponents;
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var otherProject = otherProjectMock.Object;
            var component2 = mockCodeModules[1].Object;


            var repository = TestRepository(vbe);
            project.VBComponents.Remove(component2);
            repository.Refresh(otherProject.ProjectId);
            var codeModules = repository.CodeModules(project.ProjectId).ToList();

            Assert.Contains((component2.QualifiedModuleName, component2.CodeModule), codeModules);
        }

        [Test()]
        public void CodeModulesForProjectIdDoesNotReturnRemovedCodeModulesAfterRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            projectBuilder.AddComponent("component2", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.Build();
            var mockCodeModules = projectBuilder.MockComponents;
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var component2 = mockCodeModules[1].Object;

            var repository = TestRepository(vbe);
            project.VBComponents.Remove(component2);
            repository.Refresh();
            var codeModules = repository.CodeModules(project.ProjectId).ToList();

            Assert.IsFalse(codeModules.Contains((component2.QualifiedModuleName, component2.CodeModule)));
        }

        [Test()]
        public void CodeModulesForProjectIdDoesNotReturnRemovedCodeModulesAfterRefreshForProjectIdOfContainingProject()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            projectBuilder.AddComponent("component2", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.Build();
            var mockCodeModules = projectBuilder.MockComponents;
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var component2 = mockCodeModules[1].Object;

            var repository = TestRepository(vbe);
            project.VBComponents.Remove(component2);
            repository.Refresh(project.ProjectId);
            var codeModules = repository.CodeModules(project.ProjectId).ToList();

            Assert.IsFalse(codeModules.Contains((component2.QualifiedModuleName, component2.CodeModule)));
        }

        [Test()]
        public void CodeModuleReturnsCodeModuleWithMatchingQmn()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            projectBuilder.AddComponent("component2", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.Build();
            var mockCodeModules = projectBuilder.MockCodeModules;
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var codeModule2 = mockCodeModules[1].Object;

            var repository = TestRepository(vbe);
            var returnedCodeModule = repository.CodeModule(codeModule2.QualifiedModuleName);

            Assert.AreEqual(codeModule2, returnedCodeModule);
        }

        [Test()]
        public void CodeModuleReturnsNullForUnknownQmn()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            projectBuilder.AddComponent("component2", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.Build();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;

            var repository = TestRepository(vbe);
            var newQmn = new QualifiedModuleName(String.Empty, String.Empty, "newComponent");
            var returnedCodeModule = repository.CodeModule(newQmn);

            Assert.IsNull(returnedCodeModule);
        }

        [Test()]
        public void CodeModuleReturnsNullAfterDisposal()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            projectBuilder.AddComponent("component2", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.Build();
            var mockComponents = projectBuilder.MockComponents;
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var component2 = mockComponents[1].Object;

            var repository = TestRepository(vbe);
            repository.Dispose();
            var returnedCodeModule = repository.CodeModule(component2.QualifiedModuleName);

            Assert.IsNull(returnedCodeModule);
        }

        [Test()]
        public void CodeModuleReturnsNullForQmnOfAddedCodeModuleBeforeRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.Build();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;

            var repository = TestRepository(vbe);
            var newComponent = project.VBComponents.Add(ComponentType.ClassModule);
            var returnedCodeModule = repository.CodeModule(newComponent.QualifiedModuleName);

            Assert.IsNull(returnedCodeModule);
        }

        [Test()]
        public void CodeModuleReturnsNullForQmnOfAddedCodeModuleAfterRefreshForOtherProjectId()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.Build();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var otherProject = otherProjectMock.Object;

            var repository = TestRepository(vbe);
            var newComponent = project.VBComponents.Add(ComponentType.ClassModule);
            repository.Refresh(otherProject.ProjectId);
            var returnedCodeModule = repository.CodeModule(newComponent.QualifiedModuleName);

            Assert.IsNull(returnedCodeModule);
        }

        [Test()]
        public void CodeModuleReturnsAddedCodeModuleWithMatchingQmnAfterRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.Build();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;

            var repository = TestRepository(vbe);
            var newComponent = project.VBComponents.Add(ComponentType.ClassModule);
            repository.Refresh();
            var returnedCodeModule = repository.CodeModule(newComponent.QualifiedModuleName);

            Assert.AreEqual(newComponent.CodeModule, returnedCodeModule);
        }

        [Test()]
        public void CodeModuleReturnsAddedCodeModuleWithMatchingQmnAfterRefreshForProjectIdOfContaiingProject()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.Build();
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var newComponent = project.VBComponents.Add(ComponentType.ClassModule);

            var repository = TestRepository(vbe);
            repository.Refresh(project.ProjectId);
            var returnedCodeModule = repository.CodeModule(newComponent.QualifiedModuleName);

            Assert.AreEqual(newComponent.CodeModule, returnedCodeModule);
        }

        [Test()]
        public void CodeModuleReturnsRemovedCodeModuleWithMatchingQmnBeforeRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            projectBuilder.AddComponent("component2", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.Build();
            var mockComponents = projectBuilder.MockComponents;
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var component2 = mockComponents[1].Object;

            var repository = TestRepository(vbe);
            project.VBComponents.Remove(component2);
            var returnedCodeModule = repository.CodeModule(component2.QualifiedModuleName);

            Assert.AreEqual(component2.CodeModule, returnedCodeModule);
        }

        [Test()]
        public void CodeModuleReturnsRemovedCodeModuleWithMatchingQmnAfterRefreshFroOtherProjectId()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            projectBuilder.AddComponent("component2", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.Build();
            var mockComponents = projectBuilder.MockComponents;
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var otherProject = otherProjectMock.Object;
            var component2 = mockComponents[1].Object;

            var repository = TestRepository(vbe);
            project.VBComponents.Remove(component2);
            repository.Refresh(otherProject.ProjectId);
            var returnedCodeModule = repository.CodeModule(component2.QualifiedModuleName);

            Assert.AreEqual(component2.CodeModule, returnedCodeModule);
        }

        [Test()]
        public void CodeModuleReturnsNullForQmnOfRemovedCodeModuleAfterRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            projectBuilder.AddComponent("component2", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.Build();
            var mockComponents = projectBuilder.MockComponents;
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var component2 = mockComponents[1].Object;

            var repository = TestRepository(vbe);
            project.VBComponents.Remove(component2);
            repository.Refresh();
            var returnedCodeModule = repository.CodeModule(component2.QualifiedModuleName);

            Assert.IsNull(returnedCodeModule);
        }

        [Test()]
        public void CodeModuleReturnsNullForQmnOfRemovedCodeModuleAfterRefreshForProjectIdOfContainingProject()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("project", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("component1", ComponentType.ClassModule, String.Empty);
            projectBuilder.AddComponent("component2", ComponentType.ClassModule, String.Empty);
            var projectMock = projectBuilder.Build();
            var mockComponents = projectBuilder.MockComponents;
            vbeBuilder.AddProject(projectMock);
            var otherProjectBuilder = vbeBuilder.ProjectBuilder("otherProject", ProjectProtection.Unprotected);
            otherProjectBuilder.AddComponent("otherComponent", ComponentType.ClassModule, String.Empty);
            var otherProjectMock = otherProjectBuilder.Build();
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var component2 = mockComponents[1].Object;

            var repository = TestRepository(vbe);
            project.VBComponents.Remove(component2);
            repository.Refresh(project.ProjectId);
            var returnedCodeModule = repository.CodeModule(component2.QualifiedModuleName);

            Assert.IsNull(returnedCodeModule);
        }
    }
}