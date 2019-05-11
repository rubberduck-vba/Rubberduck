using Rubberduck.Parsing.VBA.ComReferenceLoading;
using Rubberduck.Parsing.VBA.DeclarationResolving;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.Parsing.VBA.ReferenceManagement;

namespace Rubberduck.Parsing.VBA
{
    public interface IParsingStageService: 
        ICOMReferenceSynchronizer, 
        IBuiltInDeclarationLoader, 
        IParseRunner, 
        IDeclarationResolveRunner, 
        IReferenceResolveRunner, 
        IUserComProjectSynchronizer
    {
    }
}
