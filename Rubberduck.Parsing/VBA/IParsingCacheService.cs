using Rubberduck.Parsing.ComReflection;
using Rubberduck.Parsing.PreProcessing;
using Rubberduck.Parsing.VBA.DeclarationResolving;
using Rubberduck.Parsing.VBA.ReferenceManagement;

namespace Rubberduck.Parsing.VBA
{
    public interface IParsingCacheService : 
        IDeclarationFinderProvider, 
        IModuleToModuleReferenceManager, 
        IReferenceRemover, 
        ISupertypeClearer, 
        ICompilationArgumentsCache, 
        IUserComProjectRepository, 
        IProjectsToResolveFromComProjectSelector
    {
    }
}
