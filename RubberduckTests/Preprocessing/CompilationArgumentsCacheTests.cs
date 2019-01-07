using System.Collections.Generic;
using System.Linq;
using Moq;
using NUnit.Framework;
using Rubberduck.Parsing.PreProcessing;

namespace RubberduckTests.VBEditor
{
    [TestFixture()]
    public class CompilationArgumentsCacheTests
    {
        [Test]
        [Category("Preprocessor")]
        public void CompilationArgumentsCacheInitiallyReturnsNoUserConstants()
        {
            var mockArgumentsProvider = new Mock<ICompilationArgumentsProvider>();
            mockArgumentsProvider.Setup(m => m.UserDefinedCompilationArguments(It.IsAny<string>()))
                .Returns(() => new Dictionary<string, short> {{"constant", 1}});

            var argumentsProvider = mockArgumentsProvider.Object;
            var argumentsCache = new CompilationArgumentsCache(argumentsProvider);

            var userCompilationConstants = argumentsCache.UserDefinedCompilationArguments("test");

            Assert.IsTrue(userCompilationConstants.Count == 0);
        }

        [Test]
        [Category("Preprocessor")]
        public void CompilationArgumentsCacheReturnsUserConstantsAfterLoadOfTheProject()
        {
            var mockArgumentsProvider = new Mock<ICompilationArgumentsProvider>();
            mockArgumentsProvider.Setup(m => m.UserDefinedCompilationArguments(It.IsAny<string>()))
                .Returns(() => new Dictionary<string, short> { { "constant", 1 } });

            var argumentsProvider = mockArgumentsProvider.Object;
            var argumentsCache = new CompilationArgumentsCache(argumentsProvider);

            argumentsCache.ReloadCompilationArguments(new []{"test"});
            var userCompilationConstants = argumentsCache.UserDefinedCompilationArguments("test");

            var expected = 1;
            var actual = userCompilationConstants["constant"];

            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("Preprocessor")]
        public void CompilationArgumentsCacheReturnsNoUserConstantsForNotLoadedProject()
        {
            var mockArgumentsProvider = new Mock<ICompilationArgumentsProvider>();
            mockArgumentsProvider.Setup(m => m.UserDefinedCompilationArguments(It.IsAny<string>()))
                .Returns(() => new Dictionary<string, short> { { "constant", 1 } });

            var argumentsProvider = mockArgumentsProvider.Object;
            var argumentsCache = new CompilationArgumentsCache(argumentsProvider);

            argumentsCache.ReloadCompilationArguments(new[] { "test" });
            var userCompilationConstants = argumentsCache.UserDefinedCompilationArguments("notTest");

            Assert.IsTrue(userCompilationConstants.Count == 0);
        }

        [Test]
        [Category("Preprocessor")]
        public void CompilationArgumentsCacheReturnsNoUserConstantsAfterTheProjectHasBeenRemoved()
        {
            var mockArgumentsProvider = new Mock<ICompilationArgumentsProvider>();
            mockArgumentsProvider.Setup(m => m.UserDefinedCompilationArguments(It.IsAny<string>()))
                .Returns(() => new Dictionary<string, short> { { "constant", 1 } });

            var argumentsProvider = mockArgumentsProvider.Object;
            var argumentsCache = new CompilationArgumentsCache(argumentsProvider);

            argumentsCache.ReloadCompilationArguments(new[] { "test" });
            argumentsCache.RemoveCompilationArgumentsFromCache(new[] { "test" });
            var userCompilationConstants = argumentsCache.UserDefinedCompilationArguments("test");

            Assert.IsTrue(userCompilationConstants.Count == 0);
        }

        [Test]
        [Category("Preprocessor")]
        public void CompilationArgumentsCacheReturnsUserConstantsAfterLoadOfTheProjectAndRemovalOfAnother()
        {
            var mockArgumentsProvider = new Mock<ICompilationArgumentsProvider>();
            mockArgumentsProvider.Setup(m => m.UserDefinedCompilationArguments(It.IsAny<string>()))
                .Returns(() => new Dictionary<string, short> { { "constant", 1 } });

            var argumentsProvider = mockArgumentsProvider.Object;
            var argumentsCache = new CompilationArgumentsCache(argumentsProvider);

            argumentsCache.ReloadCompilationArguments(new[] { "test", "notTest" });
            argumentsCache.RemoveCompilationArgumentsFromCache(new[] { "notTest" });
            var userCompilationConstants = argumentsCache.UserDefinedCompilationArguments("test");

            var expected = 1;
            var actual = userCompilationConstants["constant"];

            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("Preprocessor")]
        public void CompilationArgumentsCacheKeepdReturningCachedUserConstantsUntilTheNextReload()
        {
            var mockArgumentsProvider = new Mock<ICompilationArgumentsProvider>();
            mockArgumentsProvider.Setup(m => m.UserDefinedCompilationArguments(It.IsAny<string>()))
                .Returns(() => new Dictionary<string, short> { { "constant", 1 } });

            var argumentsProvider = mockArgumentsProvider.Object;
            var argumentsCache = new CompilationArgumentsCache(argumentsProvider);

            argumentsCache.ReloadCompilationArguments(new[] { "test" });
            mockArgumentsProvider.Setup(m => m.UserDefinedCompilationArguments(It.IsAny<string>()))
                .Returns(() => new Dictionary<string, short> { { "constant", 2 } });
            var userCompilationConstants = argumentsCache.UserDefinedCompilationArguments("test");

            var expected = 1;
            var actual = userCompilationConstants["constant"];

            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("Preprocessor")]
        public void CompilationArgumentsCacheKeepdReturningCachedUserConstantsOnReloadOfOtherProjects()
        {
            var mockArgumentsProvider = new Mock<ICompilationArgumentsProvider>();
            mockArgumentsProvider.Setup(m => m.UserDefinedCompilationArguments(It.IsAny<string>()))
                .Returns(() => new Dictionary<string, short> { { "constant", 1 } });

            var argumentsProvider = mockArgumentsProvider.Object;
            var argumentsCache = new CompilationArgumentsCache(argumentsProvider);

            argumentsCache.ReloadCompilationArguments(new[] { "test", "notTest"});
            mockArgumentsProvider.Setup(m => m.UserDefinedCompilationArguments(It.IsAny<string>()))
                .Returns(() => new Dictionary<string, short> { { "constant", 2 } });
            argumentsCache.ReloadCompilationArguments(new[] { "notTest" });
            var userCompilationConstants = argumentsCache.UserDefinedCompilationArguments("test");

            var expected = 1;
            var actual = userCompilationConstants["constant"];

            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("Preprocessor")]
        public void CompilationArgumentsCacheUpdatesCachedUserConstantsOnReload()
        {
            var mockArgumentsProvider = new Mock<ICompilationArgumentsProvider>();
            mockArgumentsProvider.Setup(m => m.UserDefinedCompilationArguments(It.IsAny<string>()))
                .Returns(() => new Dictionary<string, short> { { "constant", 1 } });

            var argumentsProvider = mockArgumentsProvider.Object;
            var argumentsCache = new CompilationArgumentsCache(argumentsProvider);

            argumentsCache.ReloadCompilationArguments(new[] { "test" });
            mockArgumentsProvider.Setup(m => m.UserDefinedCompilationArguments(It.IsAny<string>()))
                .Returns(() => new Dictionary<string, short> { { "constant", 2 } });
            argumentsCache.ReloadCompilationArguments(new[] { "test" });
            var userCompilationConstants = argumentsCache.UserDefinedCompilationArguments("test");

            var expected = 2;
            var actual = userCompilationConstants["constant"];

            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("Preprocessor")]
        public void CompilationArgumentsCacheInitiallyReturnsEmptyCollectionOfProjectsWithChangesConstants()
        {
            var mockArgumentsProvider = new Mock<ICompilationArgumentsProvider>();
            mockArgumentsProvider.Setup(m => m.UserDefinedCompilationArguments(It.IsAny<string>()))
                .Returns(() => new Dictionary<string, short> { { "constant", 1 } });

            var argumentsProvider = mockArgumentsProvider.Object;
            var argumentsCache = new CompilationArgumentsCache(argumentsProvider);

            var projectsWithChangedConstants = argumentsCache.ProjectWhoseCompilationArgumentsChanged();

            Assert.IsTrue(projectsWithChangedConstants.Count == 0);
        }

        [Test]
        [Category("Preprocessor")]
        public void CompilationArgumentsCacheReturnsFreshlyLoadedProjectWithCompilationConstantsAsProjectWithChangedConstants()
        {
            var mockArgumentsProvider = new Mock<ICompilationArgumentsProvider>();
            mockArgumentsProvider.Setup(m => m.UserDefinedCompilationArguments(It.IsAny<string>()))
                .Returns(() => new Dictionary<string, short> { { "constant", 1 } });

            var argumentsProvider = mockArgumentsProvider.Object;
            var argumentsCache = new CompilationArgumentsCache(argumentsProvider);

            argumentsCache.ReloadCompilationArguments(new[] { "test" });
            var projectsWithChangedConstants = argumentsCache.ProjectWhoseCompilationArgumentsChanged();

            var expected = "test";
            var actual = projectsWithChangedConstants.Single();

            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("Preprocessor")]
        public void CompilationArgumentsCacheDoesNotReturnFreshlyLoadedProjectWithoutCompilationConstantsAsProjectWithChangedConstants()
        {
            var mockArgumentsProvider = new Mock<ICompilationArgumentsProvider>();
            mockArgumentsProvider.Setup(m => m.UserDefinedCompilationArguments(It.IsAny<string>()))
                .Returns(() => new Dictionary<string, short>());

            var argumentsProvider = mockArgumentsProvider.Object;
            var argumentsCache = new CompilationArgumentsCache(argumentsProvider);

            argumentsCache.ReloadCompilationArguments(new[] { "test" });
            var projectsWithChangedConstants = argumentsCache.ProjectWhoseCompilationArgumentsChanged();

            Assert.IsTrue(projectsWithChangedConstants.Count == 0);
        }

        [Test]
        [Category("Preprocessor")]
        public void CompilationArgumentsCacheDoesNotRemoveProjectsWithChangedConstantsOnReloadFromChangedCollection()
        {
            var mockArgumentsProvider = new Mock<ICompilationArgumentsProvider>();
            mockArgumentsProvider.Setup(m => m.UserDefinedCompilationArguments(It.IsAny<string>()))
                .Returns(() => new Dictionary<string, short> { { "constant", 1 } });

            var argumentsProvider = mockArgumentsProvider.Object;
            var argumentsCache = new CompilationArgumentsCache(argumentsProvider);

            argumentsCache.ReloadCompilationArguments(new[] { "test" });
            argumentsCache.ReloadCompilationArguments(new[] { "notTest" });
            var projectsWithChangedConstants = argumentsCache.ProjectWhoseCompilationArgumentsChanged();

            var expectedCount = 2;
            var actualCount = projectsWithChangedConstants.Count;

            Assert.AreEqual(expectedCount, actualCount);
            Assert.IsTrue(projectsWithChangedConstants.Contains("test"));
        }

        [Test]
        [Category("Preprocessor")]
        public void CompilationArgumentsCacheReturnsEmptyCollectionOfProjectsWithChangesConstantsAfterClearing()
        {
            var mockArgumentsProvider = new Mock<ICompilationArgumentsProvider>();
            mockArgumentsProvider.Setup(m => m.UserDefinedCompilationArguments(It.IsAny<string>()))
                .Returns(() => new Dictionary<string, short> { { "constant", 1 } });

            var argumentsProvider = mockArgumentsProvider.Object;
            var argumentsCache = new CompilationArgumentsCache(argumentsProvider);

            argumentsCache.ReloadCompilationArguments(new[] { "test" });
            argumentsCache.ClearProjectWhoseCompilationArgumentsChanged();
            var projectsWithChangedConstants = argumentsCache.ProjectWhoseCompilationArgumentsChanged();

            Assert.IsTrue(projectsWithChangedConstants.Count == 0);
        }

        [Test]
        [Category("Preprocessor")]
        public void CompilationArgumentsCacheDoesNotReturnReloadedProjectWithoutChangeOfCompilationConstantsAsProjectWithChangedConstants()
        {
            var mockArgumentsProvider = new Mock<ICompilationArgumentsProvider>();
            mockArgumentsProvider.Setup(m => m.UserDefinedCompilationArguments(It.IsAny<string>()))
                .Returns(() => new Dictionary<string, short> { { "constant", 1 } });

            var argumentsProvider = mockArgumentsProvider.Object;
            var argumentsCache = new CompilationArgumentsCache(argumentsProvider);

            argumentsCache.ReloadCompilationArguments(new[] { "test" });
            argumentsCache.ClearProjectWhoseCompilationArgumentsChanged();
            argumentsCache.ReloadCompilationArguments(new[] { "test" });
            var projectsWithChangedConstants = argumentsCache.ProjectWhoseCompilationArgumentsChanged();

            Assert.IsTrue(projectsWithChangedConstants.Count == 0);
        }

        [Test]
        [Category("Preprocessor")]
        public void CompilationArgumentsCacheReturnsReloadedProjectWithChangeOfCompilationConstantValuesAsProjectWithChangedConstants()
        {
            var mockArgumentsProvider = new Mock<ICompilationArgumentsProvider>();
            mockArgumentsProvider.Setup(m => m.UserDefinedCompilationArguments(It.IsAny<string>()))
                .Returns(() => new Dictionary<string, short> { { "constant", 1 } });

            var argumentsProvider = mockArgumentsProvider.Object;
            var argumentsCache = new CompilationArgumentsCache(argumentsProvider);

            argumentsCache.ReloadCompilationArguments(new[] { "test" });
            argumentsCache.ClearProjectWhoseCompilationArgumentsChanged();
            mockArgumentsProvider.Setup(m => m.UserDefinedCompilationArguments(It.IsAny<string>()))
                .Returns(() => new Dictionary<string, short> { { "constant", 2 } });
            argumentsCache.ReloadCompilationArguments(new[] { "test" });
            var projectsWithChangedConstants = argumentsCache.ProjectWhoseCompilationArgumentsChanged();

            var expected = "test";
            var actual = projectsWithChangedConstants.Single();

            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("Preprocessor")]
        public void CompilationArgumentsCacheReturnsReloadedProjectWithNewCompilationConstantsAsProjectWithChangedConstants()
        {
            var mockArgumentsProvider = new Mock<ICompilationArgumentsProvider>();
            mockArgumentsProvider.Setup(m => m.UserDefinedCompilationArguments(It.IsAny<string>()))
                .Returns(() => new Dictionary<string, short> { { "constant", 1 } });

            var argumentsProvider = mockArgumentsProvider.Object;
            var argumentsCache = new CompilationArgumentsCache(argumentsProvider);

            argumentsCache.ReloadCompilationArguments(new[] { "test" });
            argumentsCache.ClearProjectWhoseCompilationArgumentsChanged();
            mockArgumentsProvider.Setup(m => m.UserDefinedCompilationArguments(It.IsAny<string>()))
                .Returns(() => new Dictionary<string, short> { { "constant", 1 }, {"other", 14} });
            argumentsCache.ReloadCompilationArguments(new[] { "test" });
            var projectsWithChangedConstants = argumentsCache.ProjectWhoseCompilationArgumentsChanged();

            var expected = "test";
            var actual = projectsWithChangedConstants.Single();

            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("Preprocessor")]
        public void CompilationArgumentsCacheReturnsReloadedProjectWithRemovedCompilationConstantsAsProjectWithChangedConstants()
        {
            var mockArgumentsProvider = new Mock<ICompilationArgumentsProvider>();
            mockArgumentsProvider.Setup(m => m.UserDefinedCompilationArguments(It.IsAny<string>()))
                .Returns(() => new Dictionary<string, short> { { "constant", 1 }, { "other", 14 } });

            var argumentsProvider = mockArgumentsProvider.Object;
            var argumentsCache = new CompilationArgumentsCache(argumentsProvider);

            argumentsCache.ReloadCompilationArguments(new[] { "test" });
            argumentsCache.ClearProjectWhoseCompilationArgumentsChanged();
            mockArgumentsProvider.Setup(m => m.UserDefinedCompilationArguments(It.IsAny<string>()))
                .Returns(() => new Dictionary<string, short> { { "constant", 1 } });
            argumentsCache.ReloadCompilationArguments(new[] { "test" });
            var projectsWithChangedConstants = argumentsCache.ProjectWhoseCompilationArgumentsChanged();

            var expected = "test";
            var actual = projectsWithChangedConstants.Single();

            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("Preprocessor")]
        public void CompilationArgumentsCacheDoesNotReturnReloadedProjectWithRemovedCompilationConstantsAsProjectWithChangedConstantsMoreThenOnceIfAlreadyReported()
        {
            var mockArgumentsProvider = new Mock<ICompilationArgumentsProvider>();
            mockArgumentsProvider.Setup(m => m.UserDefinedCompilationArguments(It.IsAny<string>()))
                .Returns(() => new Dictionary<string, short> { { "constant", 1 }, { "other", 14 } });

            var argumentsProvider = mockArgumentsProvider.Object;
            var argumentsCache = new CompilationArgumentsCache(argumentsProvider);

            argumentsCache.ReloadCompilationArguments(new[] { "test" });
            mockArgumentsProvider.Setup(m => m.UserDefinedCompilationArguments(It.IsAny<string>()))
                .Returns(() => new Dictionary<string, short> { { "constant", 1 } });
            argumentsCache.ReloadCompilationArguments(new[] { "test" });
            var projectsWithChangedConstants = argumentsCache.ProjectWhoseCompilationArgumentsChanged();

            var expectedCount = 1;
            var actualCount = projectsWithChangedConstants.Count;

            Assert.AreEqual(expectedCount, actualCount);
            Assert.IsTrue(projectsWithChangedConstants.Contains("test"));
        }

        [Test]
        [Category("Preprocessor")]
        public void CompilationArgumentsCacheDoesNotReturnProjectWithChangeOfCompilationConstantsAsProjectWithChangedConstantsBeforeReload()
        {
            var mockArgumentsProvider = new Mock<ICompilationArgumentsProvider>();
            mockArgumentsProvider.Setup(m => m.UserDefinedCompilationArguments(It.IsAny<string>()))
                .Returns(() => new Dictionary<string, short> { { "constant", 1 } });

            var argumentsProvider = mockArgumentsProvider.Object;
            var argumentsCache = new CompilationArgumentsCache(argumentsProvider);

            argumentsCache.ReloadCompilationArguments(new[] { "test" });
            argumentsCache.ClearProjectWhoseCompilationArgumentsChanged();
            mockArgumentsProvider.Setup(m => m.UserDefinedCompilationArguments(It.IsAny<string>()))
                .Returns(() => new Dictionary<string, short> { { "constant", 2 } });
            var projectsWithChangedConstants = argumentsCache.ProjectWhoseCompilationArgumentsChanged();

            Assert.IsTrue(projectsWithChangedConstants.Count == 0);
        }

        [Test]
        [Category("Preprocessor")]
        public void CompilationArgumentsCacheDoesNotReturnProjectWithChangeOfCompilationConstantsAsProjectWithChangedConstantsOnReloadOfOtherProjects()
        {
            var mockArgumentsProvider = new Mock<ICompilationArgumentsProvider>();
            mockArgumentsProvider.Setup(m => m.UserDefinedCompilationArguments(It.IsAny<string>()))
                .Returns(() => new Dictionary<string, short> { { "constant", 1 } });

            var argumentsProvider = mockArgumentsProvider.Object;
            var argumentsCache = new CompilationArgumentsCache(argumentsProvider);

            argumentsCache.ReloadCompilationArguments(new[] { "test", "notTest" });
            argumentsCache.ClearProjectWhoseCompilationArgumentsChanged();
            mockArgumentsProvider.Setup(m => m.UserDefinedCompilationArguments(It.IsAny<string>()))
                .Returns(() => new Dictionary<string, short> { { "constant", 2 } });
            argumentsCache.ReloadCompilationArguments(new[] { "notTest" });
            var projectsWithChangedConstants = argumentsCache.ProjectWhoseCompilationArgumentsChanged();

            Assert.IsFalse(projectsWithChangedConstants.Contains("test"));
        }
    }
}
