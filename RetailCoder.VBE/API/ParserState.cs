using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using Rubberduck.Common;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.MenuItems;
using Rubberduck.Parsing.Preprocessing;
using System.Globalization;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.API
{
    [ComVisible(true)]
    public enum ParserState
    {
        /// <summary>
        /// Parse was requested but hasn't started yet.
        /// </summary>
        Pending,
        /// <summary>
        /// Project references are being loaded into parser state.
        /// </summary>
        LoadingReference,
        /// <summary>
        /// Code from modified modules is being parsed.
        /// </summary>
        Parsing,
        /// <summary>
        /// Parse tree is waiting to be walked for identifier resolution.
        /// </summary>
        Parsed,
        /// <summary>
        /// Resolving declarations.
        /// </summary>
        ResolvingDeclarations,
        /// <summary>
        /// Resolved declarations.
        /// </summary>
        ResolvedDeclarations,
        /// <summary>
        /// Resolving identifier references.
        /// </summary>
        ResolvingReferences,
        /// <summary>
        /// Parser state is in sync with the actual code in the VBE.
        /// </summary>
        Ready,
        /// <summary>
        /// Parsing could not be completed for one or more modules.
        /// </summary>
        Error,
        /// <summary>
        /// Parsing completed, but identifier references could not be resolved for one or more modules.
        /// </summary>
        ResolverError,
        /// <summary>
        /// This component doesn't need a state.  Use for built-in declarations.
        /// </summary>
        None,
    }

    [ComVisible(true)]
    public interface IRubberduckParserState
    {
        void Initialize(VBE vbe);

        void Parse();
        void BeginParse();

        Declaration[] AllDeclarations { get; }
        Declaration[] UserDeclarations { get; }
    }

    [ComVisible(true)]
    [Guid("3D8EAA28-8983-44D5-83AF-2EEC4C363079")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IRubberduckParserStateEvents
    {
        void OnStateChanged(ParserState state);
    }

    [ComVisible(true)]
    [Guid(ClassId)]
    [ProgId(ProgId)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ComDefaultInterface(typeof(IRubberduckParserState))]
    [ComSourceInterfaces(typeof(IRubberduckParserStateEvents))]
    [EditorBrowsable(EditorBrowsableState.Always)]
    public sealed class RubberduckParserState : IRubberduckParserState, IDisposable
    {
        private const string ClassId = "28754D11-10CC-45FD-9F6A-525A65412B7A";
        private const string ProgId = "Rubberduck.ParserState";

        private Parsing.VBA.RubberduckParserState _state;
        private AttributeParser _attributeParser;
        private RubberduckParser _parser;

        public RubberduckParserState()
        {
            UiDispatcher.Initialize();
        }

        public void Initialize(VBE vbe)
        {
            if (_parser != null)
            {
                throw new InvalidOperationException("ParserState is already initialized.");
            }

            _state = new Parsing.VBA.RubberduckParserState(vbe, new Sinks(vbe));
            _state.StateChanged += _state_StateChanged;

            Func<IVBAPreprocessor> preprocessorFactory = () => new VBAPreprocessor(double.Parse(vbe.Version, CultureInfo.InvariantCulture));
            _attributeParser = new AttributeParser(new ModuleExporter(), preprocessorFactory);
            _parser = new RubberduckParser(_state, _attributeParser, preprocessorFactory,
                new List<ICustomDeclarationLoader> { new DebugDeclarations(_state), new FormEventDeclarations(_state), new AliasDeclarations(_state) });
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

        public event Action<ParserState> OnStateChanged;

        private void _state_StateChanged(object sender, ParserStateEventArgs e)
        {
            _allDeclarations = _state.AllDeclarations
                                     .Select(item => new Declaration(item))
                                     .ToArray();
            
            _userDeclarations = _state.AllUserDeclarations
                                     .Select(item => new Declaration(item))
                                     .ToArray();

            var stateChangedHandler = OnStateChanged;
            if (stateChangedHandler != null)
            {
                UiDispatcher.Invoke(() =>
                {
                    stateChangedHandler((ParserState) Enum.Parse(typeof(ParserState), e.State.ToString()));
                });
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

            _disposed = true;
        }
    }
}
