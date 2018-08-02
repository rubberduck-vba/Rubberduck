﻿using Rubberduck.Parsing.PreProcessing;

namespace Rubberduck.Parsing.VBA
{
    public interface IParsingCacheService : IDeclarationFinderProvider, IModuleToModuleReferenceManager, IReferenceRemover, ISupertypeClearer, ICompilationArgumentsCache
    {
    }
}
