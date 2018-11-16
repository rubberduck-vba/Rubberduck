using System.Collections.Generic;
using System.IO;
using NUnit.Framework;
using Moq;
using Rubberduck.Parsing.ComReflection;
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
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SourceCodeHandling;

namespace RubberduckTests.Mocks
{
    public static class MockParser
    {
        public static RubberduckParserState ParseString(string inputCode, out QualifiedModuleName qualifiedModuleName)
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            qualifiedModuleName = new QualifiedModuleName(component);
            var parser = Create(vbe.Object);

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status == ParserState.Error)
            {
                Assert.Inconclusive("Parser Error: {0}");
            }
            return parser.State;
        }

        public static IStringParser StringParser(IVBE vbe, out ICompilationArgumentsCache compilationArgumentsCache)
        {
            var vbeVersion = double.Parse(vbe.Version, CultureInfo.InvariantCulture);
            var compilationArgumentsProvider = MockCompilationArgumentsProvider(vbeVersion);
            compilationArgumentsCache = new CompilationArgumentsCache(compilationArgumentsProvider);
            var preprocessorErrorListenerFactory = new PreprocessingParseErrorListenerFactory();
            var preprocessorParser = new VBAPreprocessorParser(preprocessorErrorListenerFactory, preprocessorErrorListenerFactory);
            var preprocessor = new VBAPreprocessor(preprocessorParser, compilationArgumentsCache);
            var mainParseErrorListenerFactory = new MainParseErrorListenerFactory();
            var mainTokenStreamParser = new VBATokenStreamParser(mainParseErrorListenerFactory, mainParseErrorListenerFactory);
            var tokenStreamProvider = new SimpleVBAModuleTokenStreamProvider();

            return new TokenStreamParserStringParserAdapterWithPreprocessing(tokenStreamProvider, mainTokenStreamParser, preprocessor);
        }

        public static (SynchronousParseCoordinator parser, IRewritingManager rewritingManager) CreateWithRewriteManager(IVBE vbe, string serializedComProjectsPath = null)
        {
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(new Mock<IVBE>());
            var declarationFinderFactory = new DeclarationFinderFactory();
            var projectRepository = new ProjectsRepository(vbe);
            var state = new RubberduckParserState(vbe, projectRepository, declarationFinderFactory, vbeEvents.Object);
            return CreateWithRewriteManager(vbe, state, projectRepository, serializedComProjectsPath);
        }

        public static SynchronousParseCoordinator Create(IVBE vbe, string serializedDeclarationsPath = null)
        {
            return CreateWithRewriteManager(vbe, serializedDeclarationsPath).parser;
        }

        public static (SynchronousParseCoordinator parser, IRewritingManager rewritingManager) CreateWithRewriteManager(IVBE vbe, RubberduckParserState state, IProjectsRepository projectRepository, string serializedComProjectsPath = null)
        {
            var path = serializedComProjectsPath ??
                       Path.Combine(Path.GetDirectoryName(Assembly.GetAssembly(typeof(MockParser)).Location), "TestFiles", "Resolver");

            var stringParser = StringParser(vbe, out var compilationsArgumentsCache);

            var projectManager = new RepositoryProjectManager(projectRepository);
            var moduleToModuleReferenceManager = new ModuleToModuleReferenceManager();
            var supertypeClearer = new SynchronousSupertypeClearer(state); 
            var parserStateManager = new SynchronousParserStateManager(state);
            var referenceRemover = new SynchronousReferenceRemover(state, moduleToModuleReferenceManager);
            var baseComDeserializer = new XmlComProjectSerializer(path);
            var comDeserializer = new StaticCachingComDeserializerDecorator(baseComDeserializer);
            var referencedDeclarationsCollector = new SerializedReferencedDeclarationsCollector(comDeserializer);
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
            var codePaneSourceCodeHandler = new CodePaneSourceCodeHandler(projectRepository);
            //We use the same handler because to achieve consistency between the return values.
            var attributesSourceCodeHandler = codePaneSourceCodeHandler;
            var moduleParser = new ModuleParser(
                codePaneSourceCodeHandler, 
                attributesSourceCodeHandler, 
                stringParser);
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
                referenceRemover);
            var parsingStageService = new ParsingStageService(
                comSynchronizer,
                builtInDeclarationLoader,
                parseRunner,
                declarationResolveRunner,
                referenceResolveRunner
                );
            var parsingCacheService = new ParsingCacheService(
                state,
                moduleToModuleReferenceManager,
                referenceRemover,
                supertypeClearer,
                compilationsArgumentsCache
                );
            var tokenStreamCache = new StateTokenStreamCache(state);
            var moduleRewriterFactory = new ModuleRewriterFactory(
                codePaneSourceCodeHandler,
                attributesSourceCodeHandler);
            var rewriterProvider = new RewriterProvider(tokenStreamCache, moduleRewriterFactory);
            var rewriteSessionFactory = new RewriteSessionFactory(state, rewriterProvider);
            var rewritingManager = new RewritingManager(rewriteSessionFactory); 

            var parser = new SynchronousParseCoordinator(
                state,
                parsingStageService,
                parsingCacheService,
                projectManager,
                parserStateManager,
                rewritingManager);

            return (parser, rewritingManager);
        }

        public static SynchronousParseCoordinator Create(IVBE vbe, RubberduckParserState state, IProjectsRepository projectRepository, string serializedComProjectsPath = null)
        {
            return CreateWithRewriteManager(vbe, state, projectRepository, serializedComProjectsPath).parser;
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

        public static (RubberduckParserState state, IRewritingManager rewritingManager) CreateAndParseWithRewritingManager(IVBE vbe, string serializedComProjectsPath = null)
        {
            var (parser, rewritingManager) = CreateWithRewriteManager(vbe, serializedComProjectsPath);
            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error)
            {
                Assert.Inconclusive("Parser Error");
            }

            return (parser.State, rewritingManager);
        }

        public static RubberduckParserState CreateAndParse(IVBE vbe, string serializedComProjectsPath = null)
        {
            return CreateAndParseWithRewritingManager(vbe, serializedComProjectsPath).state;
        }
    }
}
