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
            var codePaneSourceCodeHandler = new CodePaneSourceCodeHandler(projectRepository);
            var compilationArgumentsroviderMock = new Mock<ICompilationArgumentsProvider>();
            compilationArgumentsroviderMock.Setup(m => m.UserDefinedCompilationArguments(It.IsAny<string>()))
                .Returns(new Dictionary<string, short>());
            var compilationArgumentsProvider = compilationArgumentsroviderMock.Object;
            var attributeParser = new TestAttributeParser(() => new VBAPreprocessor(double.Parse(vbe.Version, CultureInfo.InvariantCulture), compilationArgumentsProvider), codePaneSourceCodeHandler);
            var sourceFileHandler = new Mock<ISourceFileHandler>().Object;

            var path = serializedDeclarationsPath ??
                       Path.Combine(Path.GetDirectoryName(Assembly.GetAssembly(typeof(MockParser)).Location), "TestFiles", "Resolver");
            Func<IVBAPreprocessor> preprocessorFactory = () => new VBAPreprocessor(double.Parse(vbe.Version, CultureInfo.InvariantCulture), compilationArgumentsProvider);
            var projectManager = new RepositoryProjectManager(projectRepository);
            var moduleToModuleReferenceManager = new ModuleToModuleReferenceManager();
            var supertypeClearer = new SynchronousSupertypeClearer(state); 
            var parserStateManager = new SynchronousParserStateManager(state);
            var referenceRemover = new SynchronousReferenceRemover(state, moduleToModuleReferenceManager);
            var comSynchronizer = new SynchronousCOMReferenceSynchronizer(
                state, 
                parserStateManager, 
                path);
            var builtInDeclarationLoader = new BuiltInDeclarationLoader(
                state,
                new List<ICustomDeclarationLoader>
                {
                    new DebugDeclarations(state),
                    new SpecialFormDeclarations(state),
                    new FormEventDeclarations(state),
                    new AliasDeclarations(state),
                });
            var attributesSourceCodeHanler = new SourceFileHandlerSourceCodeHandlerAdapter(sourceFileHandler, projectRepository);
            var moduleRewriterFactory = new ModuleRewriterFactory(
                codePaneSourceCodeHandler,
                attributesSourceCodeHanler);
            var parseRunner = new SynchronousParseRunner(
                state,
                parserStateManager,
                preprocessorFactory,
                attributeParser,
                codePaneSourceCodeHandler,
                attributesSourceCodeHanler,
                moduleRewriterFactory);
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
                supertypeClearer
                );

            return new SynchronousParseCoordinator(
                state,
                parsingStageService,
                parsingCacheService,
                projectManager,
                parserStateManager);
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

        // ReSharper disable once UnusedMember.Global; used by RubberduckWeb to load serialized declarations.
        public static void AddTestLibrary(this RubberduckParserState state, Stream stream)
        {
            var reader = new XmlPersistableDeclarations();
            var deserialized = reader.Load(stream);
            AddTestLibrary(state, deserialized);
        }

        private static void AddTestLibrary(RubberduckParserState state, SerializableProject deserialized)
        {
            var declarations = deserialized.Unwrap();

            foreach (var members in declarations.Where(d => d.DeclarationType != DeclarationType.Project &&
                                                            d.ParentDeclaration.DeclarationType == DeclarationType.ClassModule &&
                                                            ProceduralTypes.Contains(d.DeclarationType))
                .GroupBy(d => d.ParentDeclaration))
            {
                state.CoClasses.TryAdd(members.Select(m => m.IdentifierName).ToList(), members.First().ParentDeclaration);
            }

            foreach (var declaration in declarations)
            {
                state.AddDeclaration(declaration);
            }
        }
    }
}
