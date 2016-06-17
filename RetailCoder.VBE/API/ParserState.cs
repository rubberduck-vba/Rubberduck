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
    public interface IParserState
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
    public interface IParserStateEvents
    {
        void OnParsed();
        void OnReady();
        void OnError();
    }

    [ComVisible(true)]
    [Guid(ClassId)]
    [ProgId(ProgId)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ComDefaultInterface(typeof(IParserState))]
    [ComSourceInterfaces(typeof(IParserStateEvents))]
    [EditorBrowsable(EditorBrowsableState.Always)]
    public sealed class ParserState : IParserState, IDisposable
    {
        private const string ClassId = "28754D11-10CC-45FD-9F6A-525A65412B7A";
        private const string ProgId = "Rubberduck.ParserState";

        private readonly RubberduckParserState _state;
        private AttributeParser _attributeParser;
        private RubberduckParser _parser;

        public ParserState()
        {
            UiDispatcher.Initialize();
            _state = new RubberduckParserState();
            
            _state.StateChanged += _state_StateChanged;
        }

        public void Initialize(VBE vbe)
        {
            if (_parser != null)
            {
                throw new InvalidOperationException("ParserState is already initialized.");
            }
            Func<IVBAPreprocessor> preprocessorFactory = () => new VBAPreprocessor(double.Parse(vbe.Version, CultureInfo.InvariantCulture));
            _attributeParser = new AttributeParser(new ModuleExporter(), preprocessorFactory);
            _parser = new RubberduckParser(vbe, _state, _attributeParser, preprocessorFactory,
                new List<ICustomDeclarationLoader> { new DebugDeclarations(_state), new FormEventDeclarations(_state) });
        }

        /// <summary>
        /// Blocking call, for easier unit-test code
        /// </summary>
        public void Parse()
        {
            // blocking call
            _parser.Parse();
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

        private void _state_StateChanged(object sender, System.EventArgs e)
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

            _disposed = true;
        }
    }
}
