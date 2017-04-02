using System.Collections.Generic;
using System.Linq;
using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Refactorings.Rename
{
    public class RenameModel
    {
        private readonly IVBE _vbe;
        public IVBE VBE { get { return _vbe; } }
        
        private readonly IList<Declaration> _declarations;
        public IEnumerable<Declaration> Declarations { get { return _declarations; } }

        private Declaration _target;
        public Declaration Target
        {
            get { return _target; }
            set { _target = value; }
        }

        private readonly QualifiedSelection _selection;
        public QualifiedSelection Selection { get { return _selection; } }

        private readonly RubberduckParserState _state;
        public RubberduckParserState State { get { return _state; } }

        public string NewName { get; set; }

        private readonly IMessageBox _messageBox;

        public RenameModel(IVBE vbe, RubberduckParserState state, QualifiedSelection selection, IMessageBox messageBox)
        {
            _vbe = vbe;
            _state = state;
            _declarations = state.AllDeclarations.ToList();
            _selection = selection;
            _messageBox = messageBox;

            AcquireTarget(out _target, Selection);
        }

        private void AcquireTarget(out Declaration target, QualifiedSelection selection)
        {
            target = _declarations
                .Where(item => item.IsUserDefined && item.DeclarationType != DeclarationType.ModuleOption)
                .FirstOrDefault(item => item.IsSelected(selection) || item.References.Any(r => r.IsSelected(selection)));
        }

        public Declaration ResolveImplementationToInterfaceDeclaration(Declaration target)
        {
            if (null == target) { return target; }
            var interfaceImplementation = _declarations.FindInterfaceImplementationMembers()
                    .SingleOrDefault(m => m.Equals(target));
            if (null == interfaceImplementation )
            {
                return target;
            }

            return _declarations.FindInterfaceMember(interfaceImplementation);
        }

        public Declaration ResolveHandlerToDeclaration(Declaration handlerCandidate, DeclarationType goalType)
        {
            if (null == handlerCandidate) { return handlerCandidate; }

            if (handlerCandidate.DeclarationType != DeclarationType.Procedure
                || !(handlerCandidate.IdentifierName.Contains("_")))
            {
                return handlerCandidate;
            }

            var declarationsOfInterest = _declarations.Where(d => d.DeclarationType.HasFlag(goalType)
                    && handlerCandidate.IdentifierName.Contains(d.IdentifierName));

            if (goalType.HasFlag(DeclarationType.Control))
            {
                foreach (var controlOfInterest in declarationsOfInterest)
                {
                    var eventHandler = _declarations.FindEventHandlers(controlOfInterest)
                            .SingleOrDefault(m => m.Equals(handlerCandidate));
                    if (null != eventHandler)
                    {
                        return controlOfInterest;
                    }
                }
                return handlerCandidate;
            }

            if (goalType.HasFlag(DeclarationType.Event))
            {
                foreach( var eventOfInterest in declarationsOfInterest)
                {
                    var eventHandler = _declarations.FindHandlersForEvent(eventOfInterest)
                            .SingleOrDefault(m => m.Item2.Equals(handlerCandidate));
                    if (null != eventHandler)
                    {
                        return eventOfInterest;
                    }
                }
                return handlerCandidate;
            }

            return handlerCandidate;
        }
    }
}
