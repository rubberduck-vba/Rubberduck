using System.Collections.Concurrent;
using System.Collections.Generic;
using Rubberduck.InternalApi.Extensions;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA.Extensions;

namespace Rubberduck.Parsing.VBA.ReferenceManagement
{
    public interface IFailedResolutionStore
    {
        IReadOnlyCollection<UnboundMemberDeclaration> UnresolvedMemberDeclarations { get; }
        IReadOnlyCollection<IdentifierReference> UnboundDefaultMemberAccesses { get; }
        IReadOnlyCollection<IdentifierReference> FailedLetCoercions { get; }
        IReadOnlyCollection<IdentifierReference> FailedProcedureCoercions { get; }
        IReadOnlyCollection<IdentifierReference> FailedIndexedDefaultMemberResolutions { get; }
    }

    public interface IMutableFailedResolutionStore : IFailedResolutionStore
    {
        void AddUnresolvedMemberDeclaration(UnboundMemberDeclaration unresolvedMemberDeclaration);
        void AddUnboundDefaultMemberAccess(IdentifierReference unboundDefaultMemberAccess);
        void AddFailedLetCoercion(IdentifierReference failedLetCoercion);
        void AddFailedProcedureCoercion(IdentifierReference failedProcedureCoercion);
        void AddFailedIndexedDefaultMemberResolution(IdentifierReference failedIndexedDefaultMemberResolution);
    }

    public class FailedResolutionStore : IFailedResolutionStore
    {
        public IReadOnlyCollection<UnboundMemberDeclaration> UnresolvedMemberDeclarations { get; }
        public IReadOnlyCollection<IdentifierReference> UnboundDefaultMemberAccesses { get; }
        public IReadOnlyCollection<IdentifierReference> FailedLetCoercions { get; }
        public IReadOnlyCollection<IdentifierReference> FailedProcedureCoercions { get; }
        public IReadOnlyCollection<IdentifierReference> FailedIndexedDefaultMemberResolutions { get; }

        public FailedResolutionStore(
            IReadOnlyCollection<UnboundMemberDeclaration> unresolvedMemberDeclarations,
            IReadOnlyCollection<IdentifierReference> unboundDefaultMemberAccesses,
            IReadOnlyCollection<IdentifierReference> failedLetCoercions,
            IReadOnlyCollection<IdentifierReference> failedProcedureCoercions,
            IReadOnlyCollection<IdentifierReference> failedIndexedDefaultMemberResolutions)
        {
            UnresolvedMemberDeclarations = unresolvedMemberDeclarations;
            UnboundDefaultMemberAccesses = unboundDefaultMemberAccesses;
            FailedLetCoercions = failedLetCoercions;
            FailedProcedureCoercions = failedProcedureCoercions;
            FailedIndexedDefaultMemberResolutions = failedIndexedDefaultMemberResolutions;
        }

        public FailedResolutionStore(IMutableFailedResolutionStore mutableStore)
            :this(
                mutableStore.UnresolvedMemberDeclarations.ToHashSet(),
                mutableStore.UnboundDefaultMemberAccesses.ToHashSet(),
                mutableStore.FailedLetCoercions.ToHashSet(),
                mutableStore.FailedProcedureCoercions.ToHashSet(),
                mutableStore.FailedIndexedDefaultMemberResolutions.ToHashSet())
        {}

        public FailedResolutionStore()
        :this(
            new List<UnboundMemberDeclaration>(),
            new List<IdentifierReference>(),
            new List<IdentifierReference>(),
            new List<IdentifierReference>(),
            new List<IdentifierReference>())
        {}
    }

    public class ConcurrentFailedResolutionStore : IMutableFailedResolutionStore
    {
        private readonly ConcurrentBag<UnboundMemberDeclaration> _unresolvedMemberDeclarations = new ConcurrentBag<UnboundMemberDeclaration>();
        private readonly ConcurrentBag<IdentifierReference> _unboundDefaultMemberAccesses = new ConcurrentBag<IdentifierReference>();
        private readonly ConcurrentBag<IdentifierReference> _failedLetCoercions = new ConcurrentBag<IdentifierReference>();
        private readonly ConcurrentBag<IdentifierReference> _failedProcedureCoercions = new ConcurrentBag<IdentifierReference>();
        private readonly ConcurrentBag<IdentifierReference> _failedIndexedDefaultMemberResolutions = new ConcurrentBag<IdentifierReference>();

        public IReadOnlyCollection<UnboundMemberDeclaration> UnresolvedMemberDeclarations => _unresolvedMemberDeclarations;
        public IReadOnlyCollection<IdentifierReference> UnboundDefaultMemberAccesses => _unboundDefaultMemberAccesses;
        public IReadOnlyCollection<IdentifierReference> FailedLetCoercions => _failedLetCoercions;
        public IReadOnlyCollection<IdentifierReference> FailedProcedureCoercions => _failedProcedureCoercions;
        public IReadOnlyCollection<IdentifierReference> FailedIndexedDefaultMemberResolutions => _failedIndexedDefaultMemberResolutions;

        public void AddUnresolvedMemberDeclaration(UnboundMemberDeclaration unresolvedMemberDeclaration)
        {
            _unresolvedMemberDeclarations.Add(unresolvedMemberDeclaration);
        }

        public void AddUnboundDefaultMemberAccess(IdentifierReference unboundDefaultMemberAccess)
        {
            _unboundDefaultMemberAccesses.Add(unboundDefaultMemberAccess);
        }

        public void AddFailedLetCoercion(IdentifierReference failedLetCoercion)
        {
            _failedLetCoercions.Add(failedLetCoercion);
        }

        public void AddFailedProcedureCoercion(IdentifierReference failedProcedureCoercion)
        {
            _failedProcedureCoercions.Add(failedProcedureCoercion);
        }

        public void AddFailedIndexedDefaultMemberResolution(IdentifierReference failedIndexedDefaultMemberResolution)
        {
            _failedIndexedDefaultMemberResolutions.Add(failedIndexedDefaultMemberResolution);
        }
    }
}