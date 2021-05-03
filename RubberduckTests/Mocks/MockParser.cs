using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using NUnit.Framework;
using Moq;
using Rubberduck.Parsing.ComReflection;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.Symbols.DeclarationLoaders;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using System.Globalization;
using System.Reflection;
using System.Threading;
using Rubberduck.Parsing.PreProcessing;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA.ComReferenceLoading;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Parsing.VBA.DeclarationResolving;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.Parsing.VBA.Parsing.ParsingExceptions;
using Rubberduck.Parsing.VBA.ReferenceManagement;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SourceCodeHandling;
using Rubberduck.VBEditor.Utility;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Settings;
using Rubberduck.SettingsProvider;

namespace RubberduckTests.Mocks
{
    public static class MockParser
    {
        public static RubberduckParserState ParseString(string inputCode, out QualifiedModuleName qualifiedModuleName)
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out IVBComponent component);
            qualifiedModuleName = new QualifiedModuleName(component);
            var parser = Create(vbe.Object);

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status == ParserState.Error)
            {
                Assert.Inconclusive("Parser Error: {0}");
            }
            return parser.State;
        }

        public static (SynchronousParseCoordinator parser, IRewritingManager rewritingManager) CreateWithRewriteManager(IVBE vbe, string serializedComProjectsPath = null, Mock<IVbeEvents> vbeEvents = null, IDictionary<string, IEnumerable<string>> documentModuleSupertypeNames = null)
        {
            var declarationFinderFactory = new DeclarationFinderFactory();
            var projectRepository = new ProjectsRepository(vbe);
            var state = new RubberduckParserState(vbe, projectRepository, declarationFinderFactory, vbeEvents?.Object ?? MockVbeEvents.CreateMockVbeEvents(new Mock<IVBE>()).Object);
            return CreateWithRewriteManager(vbe, state, projectRepository, serializedComProjectsPath, documentModuleSupertypeNames);
        }

        public static SynchronousParseCoordinator Create(IVBE vbe, string serializedDeclarationsPath = null, Mock<IVbeEvents> vbeEvents = null, IDictionary<string, IEnumerable<string>> documentModuleSupertypeNames = null)
        {
            return CreateWithRewriteManager(vbe, serializedDeclarationsPath, vbeEvents, documentModuleSupertypeNames).parser;
        }

        public static (SynchronousParseCoordinator parser, IRewritingManager rewritingManager) CreateWithRewriteManager(IVBE vbe, RubberduckParserState state, IProjectsRepository projectRepository, string serializedComProjectsPath = null, IDictionary<string, IEnumerable<string>> documentModuleSupertypeNames = null)
        {
            var vbeVersion = double.Parse(vbe.Version, CultureInfo.InvariantCulture);
            var compilationArgumentsProvider = MockCompilationArgumentsProvider(vbeVersion);
            var compilationsArgumentsCache = new CompilationArgumentsCache(compilationArgumentsProvider);
            var userComProjectsRepository = MockUserComProjectRepository();
            var documentSuperTypesProvider = MockDocumentSuperTypeNamesProvider(documentModuleSupertypeNames);
            var ignoredProjectsSettingsProvider = MockIgnoredProjectsSettingsProviderMock();
            var projectsToBeLoadedFromComSelector = new ProjectsToResolveFromComProjectsSelector(projectRepository, ignoredProjectsSettingsProvider);

            var path = serializedComProjectsPath ??
                       Path.Combine(Path.GetDirectoryName(Assembly.GetAssembly(typeof(MockParser)).Location), "Testfiles", "Resolver");
            var preprocessorErrorListenerFactory = new PreprocessingParseErrorListenerFactory();
            var preprocessorParser = new VBAPreprocessorParser(preprocessorErrorListenerFactory, preprocessorErrorListenerFactory);
            var preprocessor = new VBAPreprocessor(preprocessorParser, compilationsArgumentsCache);
            var mainParseErrorListenerFactory = new MainParseErrorListenerFactory();
            var mainTokenStreamParser = new VBATokenStreamParser(mainParseErrorListenerFactory, mainParseErrorListenerFactory);
            var tokenStreamProvider = new SimpleVBAModuleTokenStreamProvider();
            var stringParser = new TokenStreamParserStringParserAdapterWithPreprocessing(tokenStreamProvider, mainTokenStreamParser, preprocessor);
            var vbaParserAnnotationFactory = new VBAParserAnnotationFactory(WellKnownAnnotations());
            var projectManager = new RepositoryProjectManager(projectRepository);
            var moduleToModuleReferenceManager = new ModuleToModuleReferenceManager();
            var supertypeClearer = new SynchronousSupertypeClearer(state); 
            var parserStateManager = new SynchronousParserStateManager(state);
            var referenceRemover = new SynchronousReferenceRemover(state, moduleToModuleReferenceManager);
            var baseComDeserializer = new XmlComProjectSerializer(new MockFileSystem(), path);
            var comDeserializer = new StaticCachingComDeserializerDecorator(baseComDeserializer);
            var declarationsFromComProjectLoader = new DeclarationsFromComProjectLoader();
            var referencedDeclarationsCollector = new SerializedReferencedDeclarationsCollector(declarationsFromComProjectLoader, comDeserializer);
            var userComProjectSynchronizer = new UserComProjectSynchronizer(state, declarationsFromComProjectLoader, userComProjectsRepository, projectsToBeLoadedFromComSelector);
            var comSynchronizer = new SynchronousCOMReferenceSynchronizer(
                state, 
                parserStateManager,
                projectRepository,
                referencedDeclarationsCollector);
            var builtInDeclarationLoader = new BuiltInDeclarationLoader(
                state,
                new List<ICustomDeclarationLoader>
                {
                    new DebugDeclarations(state),
                    new SpecialFormDeclarations(state),
                    new FormEventDeclarations(state),
                    new AliasDeclarations(state),
                });
            var codePaneComponentSourceCodeHandler = new CodeModuleComponentSourceCodeHandler();
            var codePaneSourceCodeHandler = new ComponentSourceCodeHandlerSourceCodeHandlerAdapter(codePaneComponentSourceCodeHandler, projectRepository);
            //We use the same handler because to achieve consistency between the return values.
            var attributesSourceCodeHandler = codePaneSourceCodeHandler;
            var moduleParser = new ModuleParser(
                codePaneSourceCodeHandler, 
                attributesSourceCodeHandler, 
                stringParser,
                vbaParserAnnotationFactory);
            var parseRunner = new SynchronousParseRunner(
                state,
                parserStateManager,
                moduleParser);
            var declarationResolveRunner = new SynchronousDeclarationResolveRunner(
                state, 
                parserStateManager, 
                comSynchronizer);
            var referenceResolveRunner = new SynchronousReferenceResolveRunner(
                state,
                parserStateManager,
                moduleToModuleReferenceManager,
                referenceRemover,
                documentSuperTypesProvider);
            var parsingStageService = new ParsingStageService(
                comSynchronizer,
                builtInDeclarationLoader,
                parseRunner,
                declarationResolveRunner,
                referenceResolveRunner,
                userComProjectSynchronizer
                );
            var parsingCacheService = new ParsingCacheService(
                state,
                moduleToModuleReferenceManager,
                referenceRemover,
                supertypeClearer,
                compilationsArgumentsCache,
                userComProjectsRepository,
                projectsToBeLoadedFromComSelector
                );
            var tokenStreamCache = new StateTokenStreamCache(state);
            var moduleRewriterFactory = new ModuleRewriterFactory(
                codePaneSourceCodeHandler,
                attributesSourceCodeHandler);
            var rewriterProvider = new RewriterProvider(tokenStreamCache, moduleRewriterFactory);
            var selectionService = new SelectionService(vbe, projectRepository);
            var selectionRecoverer = new SelectionRecoverer(selectionService, state);
            var rewriteSessionFactory = new RewriteSessionFactory(state, rewriterProvider, selectionRecoverer);
            var stubMembersAttributeRecoverer = new Mock<IMemberAttributeRecovererWithSettableRewritingManager>().Object;
            var rewritingManager = new RewritingManager(rewriteSessionFactory, stubMembersAttributeRecoverer); 

            var parser = new SynchronousParseCoordinator(
                state,
                parsingStageService,
                parsingCacheService,
                projectManager,
                parserStateManager,
                rewritingManager);

            return (parser, rewritingManager);
        }

        public static IEnumerable<IAnnotation> WellKnownAnnotations()
        {
            return Assembly.GetAssembly(typeof(IAnnotation))
                .GetTypes()
                .Where(candidate => typeof(IAnnotation).IsAssignableFrom(candidate)
                    && !candidate.IsAbstract)
                .Select(t => (IAnnotation)Activator.CreateInstance(t));
        }

        public static SynchronousParseCoordinator Create(IVBE vbe, RubberduckParserState state, IProjectsRepository projectRepository, string serializedComProjectsPath = null)
        {
            return CreateWithRewriteManager(vbe, state, projectRepository, serializedComProjectsPath).parser;
        }

        private static IDocumentModuleSuperTypeNamesProvider MockDocumentSuperTypeNamesProvider(IDictionary<string, IEnumerable<string>> documentModuleSupertypeNames = null)
        {
            var mock = new Mock<IDocumentModuleSuperTypeNamesProvider>();
            if (documentModuleSupertypeNames == null)
            {
                return mock.Object;
            }

            mock.Setup(m => m.GetSuperTypeNamesFor(It.IsAny<DocumentModuleDeclaration>()))
                .Returns<DocumentModuleDeclaration>(declaration =>
                {
                    if (documentModuleSupertypeNames.TryGetValue(declaration.IdentifierName, out var superTypeNames))
                    {
                        return superTypeNames;
                    }
                    else
                    {
                        return Enumerable.Empty<string>();
                    }
                });

            return mock.Object;
        }

        private static IUserComProjectRepository MockUserComProjectRepository()
        {
            var userComProjectsRepository = new Mock<IUserComProjectRepository>();
            userComProjectsRepository.Setup(m => m.UserProject(It.IsAny<string>())).Returns((string projectId) => null);
            userComProjectsRepository.Setup(m => m.UserProjects()).Returns(() => null);
            return userComProjectsRepository.Object;
        }

        private static IConfigurationService<IgnoredProjectsSettings> MockIgnoredProjectsSettingsProviderMock()
        {
            var mock = new Mock<IConfigurationService<IgnoredProjectsSettings>>();
            var defaultSettings = new IgnoredProjectsSettings();
            var currentSettings = new IgnoredProjectsSettings();
            mock.Setup(m => m.Read()).Returns(currentSettings);
            mock.Setup(m => m.ReadDefaults()).Returns(defaultSettings);
            return mock.Object;
        }

        private static ICompilationArgumentsProvider MockCompilationArgumentsProvider(double vbeVersion)
        {
            var compilationArgumentsProviderMock = new Mock<ICompilationArgumentsProvider>();
            var predefinedCompilationConstants = new VBAPredefinedCompilationConstants(vbeVersion);
            compilationArgumentsProviderMock.Setup(m => m.UserDefinedCompilationArguments(It.IsAny<string>()))
                .Returns(new Dictionary<string, short>());
            compilationArgumentsProviderMock.Setup(m => m.PredefinedCompilationConstants)
                .Returns(() => predefinedCompilationConstants);
            var compilationArgumentsProvider = compilationArgumentsProviderMock.Object;
            return compilationArgumentsProvider;
        }

        public static (RubberduckParserState state, IRewritingManager rewritingManager) CreateAndParseWithRewritingManager(IVBE vbe, string serializedComProjectsPath = null, IDictionary<string, IEnumerable<string>> documentModuleSupertypeNames = null)
        {
            var (parser, rewritingManager) = CreateWithRewriteManager(vbe, serializedComProjectsPath, documentModuleSupertypeNames:documentModuleSupertypeNames);
            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error)
            {
                Assert.Inconclusive("Parser Error");
            }

            return (parser.State, rewritingManager);
        }

        public static RubberduckParserState CreateAndParse(IVBE vbe, string serializedComProjectsPath = null, IDictionary<string, IEnumerable<string>> documentModuleSupertypeNames = null) => 
            CreateAndParseWithRewritingManager(vbe, serializedComProjectsPath, documentModuleSupertypeNames:documentModuleSupertypeNames).state;
        
    }
}
