﻿using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Threading;

namespace Rubberduck.Parsing.VBA
{
    public abstract class ParserStateManagerBase:IParserStateManager 
    {
        protected readonly RubberduckParserState _state;

        protected ParserStateManagerBase(RubberduckParserState state)
        {
            _state = state ?? throw new ArgumentNullException(nameof(state));
        }


        public abstract void SetModuleStates(IReadOnlyCollection<QualifiedModuleName> modules, ParserState parserState, CancellationToken token, bool evaluateOverallParserState = true);


        public ParserState OverallParserState => _state.Status;

        public ParserState GetModuleState(QualifiedModuleName module)
        {
            return _state.GetModuleState(module);
        }

        public void EvaluateOverallParserState(CancellationToken token)
        {
            _state.EvaluateParserState();
        }

        public void SetModuleState(QualifiedModuleName module, ParserState parserState, CancellationToken token, bool evaluateOverallParserState = true)
        {
            _state.SetModuleState(module, parserState, token, null, evaluateOverallParserState);
        }

        public void SetStatusAndFireStateChanged(object requestor, ParserState status, CancellationToken token)
        {
            _state.SetStatusAndFireStateChanged(requestor, status);
        }
    }
}
