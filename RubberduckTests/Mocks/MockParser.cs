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
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SourceCodeHandling;

namespace RubberduckTests.Mocks
{
    public static class MockParser
    {
        public static RubberduckParserState ParseString(string inputCode, out QualifiedModuleName qualifiedModuleName)
        {
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            qualifiedModuleName = new QualifiedModuleName(component);
            var parser = Create(vbe.Object);

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status == ParserState.Error)
            {
                Assert.Inconclusive("Parser Error: {0}");
            }
            return parser.State;
        }

        public static SynchronousParseCoordinator Create(IVBE vbe, string serializedDeclarationsPath = null)
        {
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(new Moq.Mock<IVBE>());
            var declarationFinderFactory = new DeclarationFinderFactory();
            var projectRepository = new ProjectsRepository(vbe);
            var state = new RubberduckParserState(vbe, projectRepository, declarationFinderFactory, vbeEvents.Object);
            return Create(vbe, state, projectRepository, serializedDeclarationsPath);
        }

        public static SynchronousParseCoordinator Create(IVBE vbe, RubberduckParserState state, IProjectsRepository projectRepository, string serializedDeclarationsPath = null)
        {
            var vbeVersion = double.Parse(vbe.Version, CultureInfo.InvariantCulture);
            var compilationArgumentsProvider = MockCompilationArgumentsProvider(vbeVersion);
            var compilationsArgumentsCache = new CompilationArgumentsCache(compilationArgumentsProvider);

            var path = serializedDeclarationsPath ??
                       Path.Combine(Path.GetDirectoryName(Assembly.GetAssembly(typeof(MockParser)).Location), "TestFiles", "Resolver");
            var preprocessorErrorListenerFactory = new PreprocessingParseErrorListenerFactory();
            var preprocessorParser = new VBAPreprocessorParser(preprocessorErrorListenerFactory, preprocessorErrorListenerFactory);
            var preprocessor = new VBAPreprocessor(preprocessorParser, compilationsArgumentsCache);
            var mainParseErrorListenerFactory = new MainParseErrorListenerFactory();
            var mainTokenStreamParser = new VBATokenStreamParser(mainParseErrorListenerFactory, mainParseErrorListenerFactory);
            var tokenStreamProvider = new SimpleVBAModuleTokenStreamProvider();
            var stringParser = new TokenStreamParserStringParserAdapterWithPreprocessing(tokenStreamProvider, mainTokenStreamParser, preprocessor);
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
                }); var codePaneSourceCodeHandler = new CodePaneSourceCodeHandler(projectRepository);
            //We use the same handler because to achieve consistency between the return values.
            var attributesSourceCodeHandler = codePaneSourceCodeHandler;
            var moduleRewriterFactory = new ModuleRewriterFactory(
                codePaneSourceCodeHandler,
                attributesSourceCodeHandler);
            var moduleParser = new ModuleParser(
                codePaneSourceCodeHandler, 
                attributesSourceCodeHandler, 
                stringParser, 
                moduleRewriterFactory);
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

            return new SynchronousParseCoordinator(
                state,
                parsingStageService,
                parsingCacheService,
                projectManager,
                parserStateManager);
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

        public static RubberduckParserState CreateAndParse(IVBE vbe, string serializedDeclarationsPath = null)
        {
            var parser = Create(vbe, serializedDeclarationsPath);
            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            return parser.State;
        }

        private static readonly HashSet<DeclarationType> ProceduralTypes =
            new HashSet<DeclarationType>(new[]
            {
                DeclarationType.Procedure, DeclarationType.Function, DeclarationType.PropertyGet,
                DeclarationType.PropertyLet, DeclarationType.PropertySet
            });
    }
}
