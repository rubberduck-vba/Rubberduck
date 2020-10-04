using System;
using System.Collections.Generic;
using System.Linq;
using NLog;
using Rubberduck.InternalApi.Extensions;
using Rubberduck.Parsing.VBA.Extensions;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;

namespace Rubberduck.Parsing.Rewriter
{
    public abstract class RewriteSessionBase : IExecutableRewriteSession
    {
        protected readonly IDictionary<QualifiedModuleName, IExecutableModuleRewriter> CheckedOutModuleRewriters = new Dictionary<QualifiedModuleName, IExecutableModuleRewriter>();
        protected readonly IRewriterProvider RewriterProvider;

        protected readonly ISelectionRecoverer SelectionRecoverer;
        private readonly Func<IRewriteSession, bool> _rewritingAllowed;

        protected readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private readonly object _statusLockObject = new object();

        public abstract CodeKind TargetCodeKind { get; }

        protected RewriteSessionBase(IRewriterProvider rewriterProvider, ISelectionRecoverer selectionRecoverer, Func<IRewriteSession, bool> rewritingAllowed)
        {
            RewriterProvider = rewriterProvider;
            _rewritingAllowed = rewritingAllowed;
            SelectionRecoverer = selectionRecoverer;
        }

        public IReadOnlyCollection<QualifiedModuleName> CheckedOutModules => CheckedOutModuleRewriters.Keys.ToHashSet();

        private RewriteSessionState _status = RewriteSessionState.Valid;
        public RewriteSessionState Status
        {
            get
            {
                lock (_statusLockObject)
                {
                    return _status;
                }
            }
            set
            {
                lock (_statusLockObject)
                {
                    if (_status == RewriteSessionState.Valid)
                    {
                        _status = value;
                    }
                }
            }
        }

        public IModuleRewriter CheckOutModuleRewriter(QualifiedModuleName module)
        {
            if (CheckedOutModuleRewriters.TryGetValue(module, out var rewriter))
            {
                return rewriter;
            }
            
            rewriter = ModuleRewriter(module);
            CheckedOutModuleRewriters.Add(module, rewriter);

            if (rewriter.IsDirty)
            {
                //The parse tree is stale.
                Status = RewriteSessionState.StaleParseTree;
            }

            return rewriter;
        }

        protected abstract IExecutableModuleRewriter ModuleRewriter(QualifiedModuleName module);

        public bool TryRewrite()
        {
            if (!CheckedOutModuleRewriters.Any())
            {
                return true;
            }

            //This is thread-safe because, once invalidated, there is no way back.
            if (Status != RewriteSessionState.Valid)
            {
                Logger.Warn($"Tried to execute Rewrite on a RewriteSession that was in the invalid status {Status}.");
                return false;
            }            

            if (!_rewritingAllowed(this))
            {
                Logger.Debug("Tried to execute Rewrite on a RewriteSession when rewriting was no longer allowed.");
                return false;
            }

            PrimeSelectionRecovery();

            return TryRewriteInternal();
        }

        protected abstract bool TryRewriteInternal();

        private void PrimeSelectionRecovery()
        {
            SelectionRecoverer.SaveSelections(CheckedOutModules);

            foreach (var (module, rewriter) in CheckedOutModuleRewriters)
            {
                if (rewriter.SelectionOffset.HasValue)
                {
                    SelectionRecoverer.AdjustSavedSelection(module, rewriter.SelectionOffset.Value);
                }
                if (rewriter.Selection.HasValue)
                {
                    SelectionRecoverer.ReplaceSavedSelection(module, rewriter.Selection.Value);
                }
            }

            SelectionRecoverer.RecoverSavedSelectionsOnNextParse();
        }
    }
}