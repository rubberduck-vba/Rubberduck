using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace Rubberduck.Parsing.VBA
{
    public class ParserStateChangeCallbackManager : IParserStateChangeCallbackManager
    {
        private readonly Dictionary<ParserState, ConcurrentDictionary<Action<CancellationToken>, byte>> _callbacks =
            new Dictionary<ParserState, ConcurrentDictionary<Action<CancellationToken>, byte>>();

        public ParserStateChangeCallbackManager()
        {
            foreach (ParserState value in Enum.GetValues(typeof(ParserState)))
            {
                _callbacks.Add(value, new ConcurrentDictionary<Action<CancellationToken>, byte>());
            }
        }

        public void RegisterCallback(Action<CancellationToken> callback, ParserState state)
        {
            foreach (ParserState value in Enum.GetValues(typeof(ParserState)))
            {
                if (!state.HasFlag(value)) { continue; }

                ConcurrentDictionary<Action<CancellationToken>, byte> callbacks;
                if (!_callbacks.ContainsKey(value))
                {
                    lock (_callbacks)
                    {
                        callbacks = new ConcurrentDictionary<Action<CancellationToken>, byte>();
                        _callbacks.Add(value, callbacks);
                    }
                }
                else
                {
                    callbacks = _callbacks[value];
                }

                callbacks.TryAdd(callback, 0);
            }
        }

        public void UnregisterCallback(Action<CancellationToken> callback)
        {
            foreach (var value in _callbacks.Values)
            {
                if (value.ContainsKey(callback))
                {
                    byte b;
                    value.TryRemove(callback, out b);
                }
            }
        }

        public void RunCallbacks(ParserState state, CancellationToken token)
        {
            foreach (var callback in _callbacks[state].Keys)
            {
                if (token.IsCancellationRequested) { break; }

                Task.Run(() => callback(token), token);
            }
        }
    }
}