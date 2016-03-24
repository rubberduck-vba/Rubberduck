using System;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using Rubberduck.Common;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.API
{
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IParserState
    {
        void Initialize(VBE vbe);

        void Parse();
        void BeginParse();

        Declaration[] AllDeclarations 
        {
            get; 
        }

        Declaration[] UserDeclarations
        {
            get;
        }
    }

    [ComVisible(true)]
    [Guid(ClassId)]
    [ProgId(ProgId)]
    [ComDefaultInterface(typeof(IParserState))]
    [EditorBrowsable(EditorBrowsableState.Always)]
    public class ParserState : IParserState
    {
        private const string ClassId = "28754D11-10CC-45FD-9F6A-525A65412B7A";
        private const string ProgId = "Rubberduck.ParserState";

        private readonly RubberduckParserState _state;
        private readonly AttributeParser _attributeParser;

        private RubberduckParser _parser;

        public ParserState()
        {
            _state = new RubberduckParserState();
            _attributeParser = new AttributeParser(new ModuleExporter());
            
            _state.StateChanged += _state_StateChanged;
        }

        public void Initialize(VBE vbe)
        {
            if (_parser != null)
            {
                throw new InvalidOperationException("ParserState is already initialized.");
            }

            _parser = new RubberduckParser(vbe, _state, _attributeParser);
        }

        public void Parse()
        {
            // blocking call
            _parser.Parse();
        }

        public void BeginParse()
        {
            // non-blocking call
            _state.OnParseRequested(this);
        }

        private void _state_StateChanged(object sender, System.EventArgs e)
        {
            _allDeclarations = _state.AllDeclarations
                                     .Select(item => new Declaration(item))
                                     .ToArray();
            
            _userDeclarations = _state.AllUserDeclarations
                                     .Select(item => new Declaration(item))
                                     .ToArray();
        }

        private Declaration[] _allDeclarations;

        public Declaration[] AllDeclarations
        {
            [return: MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)]
            get { return _allDeclarations; }
        }

        private Declaration[] _userDeclarations;
        public Declaration[] UserDeclarations
        {
            [return: MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)]
            get { return _userDeclarations; }
        }
    }
}
