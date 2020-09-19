using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings
{
    public interface IEncapsulateFieldCollectionsProviderFactory
    {
        IEncapsulateFieldCollectionsProvider Create(QualifiedModuleName qualifiedModuleName);
    }

    public class EncapsulateFieldCollectionsProviderFactory : IEncapsulateFieldCollectionsProviderFactory
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IEncapsulateFieldCandidateFactory _encapsulateFieldCandidateFactory;
        private readonly IObjectStateUserDefinedTypeFactory _objectStateUserDefinedTypeFactory;

        public EncapsulateFieldCollectionsProviderFactory(
            IDeclarationFinderProvider declarationFinderProvider,
            IEncapsulateFieldCandidateFactory encapsulateFieldCandidateFactory,
            IObjectStateUserDefinedTypeFactory objectStateUserDefinedTypeFactory)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _encapsulateFieldCandidateFactory = encapsulateFieldCandidateFactory;
            _objectStateUserDefinedTypeFactory = objectStateUserDefinedTypeFactory;
        }

        public IEncapsulateFieldCollectionsProvider Create(QualifiedModuleName qualifiedModuleName)
        {
            return new EncapsulateFieldCollectionsProvider(
                _declarationFinderProvider,
                _encapsulateFieldCandidateFactory,
                _objectStateUserDefinedTypeFactory,
                qualifiedModuleName);
        }
    }

}
