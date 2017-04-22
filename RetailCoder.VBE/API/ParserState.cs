using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.Common;
using Rubberduck.Parsing.Symbols.DeclarationLoaders;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.MenuItems;
using System.Globalization;
using Rubberduck.Parsing.PreProcessing;
using Rubberduck.VBEditor.SafeComWrappers.VBA;

namespace Rubberduck.API
{
    [ComVisible(true)]
    public interface IParserState
    {
        void Initialize(Microsoft.Vbe.Interop.VBE vbe);

        void Parse();
        void BeginParse();

        Declaration[] AllDeclarations { get; }
        Declaration[] UserDeclarations { get; }
    }

    [ComVisible(true)]
    [Guid(RubberduckGuid.IParserStateEventsGuid)]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IParserStateEvents
    {
        void OnParsed();
        void OnReady();
        void OnError();
    }

    [ComVisible(true)]
    [Guid(RubberduckGuid.ParserStateClassGuid)]
    [ProgId(RubberduckProgId.ParserStateProgId)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ComDefaultInterface(typeof(IParserState))]
    [ComSourceInterfaces(typeof(IParserStateEvents))]
    [EditorBrowsable(EditorBrowsableState.Always)]
    public sealed class ParserState : IParserState, IDisposable
    {
        private RubberduckParserState _state;
        private AttributeParser _attributeParser;
        private ParseCoordinator _parser;
        private VBE _vbe;

        public ParserState()
        {
            UiDispatcher.Initialize();
        }

        public void Initialize(Microsoft.Vbe.Interop.VBE vbe)
        {
            if (_parser != null)
            {
                throw new InvalidOperationException("ParserState is already initialized.");
            }

            _vbe = new VBE(vbe);
            _state = new RubberduckParserState(null);
            _state.StateChanged += _state_StateChanged;

            Func<IVBAPreprocessor> preprocessorFactory = () => new VBAPreprocessor(double.Parse(_vbe.Version, CultureInfo.InvariantCulture));
            _attributeParser = new AttributeParser(new ModuleExporter(), preprocessorFactory);
            var projectManager = new ProjectManager(_state, _vbe);
            var moduleToModuleReferenceManager = new ModuleToModuleReferenceManager();
            var parserStateManager = new ParserStateManager(_state);
            var referenceRemover = new ReferenceRemover(_state, moduleToModuleReferenceManager);
            var comSynchronizer = new COMReferenceSynchronizer(_state, parserStateManager);
            var builtInDeclarationLoader = new BuiltInDeclarationLoader(
                _state,
                new List<ICustomDeclarationLoader>
                    {
                        new DebugDeclarations(_state),
                        new SpecialFormDeclarations(_state),
                        new FormEventDeclarations(_state),
                        new AliasDeclarations(_state),
                        //new RubberduckApiDeclarations(_state)
                    }
                );
            var parseRunner = new ParseRunner(
                _state,
                parserStateManager,
                preprocessorFactory,
                _attributeParser);
            var declarationResolveRunner = new DeclarationResolveRunner(
                _state, 
                parserStateManager, 
                comSynchronizer);
            var referenceResolveRunner = new ReferenceResolveRunner(
                _state,
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

            _parser = new ParseCoordinator(
                _state,
                parsingStageService,
                projectManager,
                moduleToModuleReferenceManager,
                parserStateManager,
                referenceRemover);
        }

        /// <summary>
        /// Blocking call, for easier unit-test code
        /// </summary>
        public void Parse()
        {
            // blocking call
            _parser.Parse(new System.Threading.CancellationTokenSource());
        }

        /// <summary>
        /// Begins asynchronous parsing
        /// </summary>
        public void BeginParse()
        {
            // non-blocking call
            UiDispatcher.Invoke(() => _state.OnParseRequested(this));
        }

        public event Action OnParsed;
        public event Action OnReady;
        public event Action OnError;

        private void _state_StateChanged(object sender, EventArgs e)
        {
            _allDeclarations = _state.AllDeclarations
                                     .Select(item => new Declaration(item))
                                     .ToArray();
            
            _userDeclarations = _state.AllUserDeclarations
                                     .Select(item => new Declaration(item))
                                     .ToArray();

            var errorHandler = OnError;
            if (_state.Status == Parsing.VBA.ParserState.Error && errorHandler != null)
            {
                UiDispatcher.Invoke(errorHandler.Invoke);
            }

            var parsedHandler = OnParsed;
            if (_state.Status == Parsing.VBA.ParserState.Parsed && parsedHandler != null)
            {
                UiDispatcher.Invoke(parsedHandler.Invoke);
            }

            var readyHandler = OnReady;
            if (_state.Status == Parsing.VBA.ParserState.Ready && readyHandler != null)
            {
                UiDispatcher.Invoke(readyHandler.Invoke);
            }
        }

        private Declaration[] _allDeclarations;

        public Declaration[] AllDeclarations
        {
            //[return: MarshalAs(UnmanagedType.SafeArray/*, SafeArraySubType = VarEnum.VT_VARIANT*/)]
            get { return _allDeclarations; }
        }

        private Declaration[] _userDeclarations;
        public Declaration[] UserDeclarations
        {
            //[return: MarshalAs(UnmanagedType.SafeArray/*, SafeArraySubType = VarEnum.VT_VARIANT*/)]
            get { return _userDeclarations; }
        }

        private bool _disposed;
        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            if (_state != null)
            {
                _state.StateChanged -= _state_StateChanged;
            }


            //_vbe.Release();            
            _disposed = true;
        }
    }
}
