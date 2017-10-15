﻿using System;
using System.Collections.Generic;
using System.Threading;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.VBA
{
    public class ParsingCacheService : IParsingCacheService
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IModuleToModuleReferenceManager _moduleToModuleReferenceManager;
        private readonly IReferenceRemover _referenceRemover;
        private readonly ISupertypeClearer _supertypeClearer;

        public ParsingCacheService(
            IDeclarationFinderProvider declarationFinderProvider,
            IModuleToModuleReferenceManager moduleToModuleReferenceManager,
            IReferenceRemover referenceRemover,
            ISupertypeClearer supertypeClearer)
        {
            if(declarationFinderProvider == null)
            {
                throw new ArgumentNullException(nameof(declarationFinderProvider));
            }
            if (moduleToModuleReferenceManager == null)
            {
                throw new ArgumentNullException(nameof(moduleToModuleReferenceManager));
            }
            if (referenceRemover == null)
            {
                throw new ArgumentNullException(nameof(referenceRemover));
            }
            if (supertypeClearer == null)
            {
                throw new ArgumentNullException(nameof(supertypeClearer));
            }
            _declarationFinderProvider = declarationFinderProvider;
            _moduleToModuleReferenceManager = moduleToModuleReferenceManager;
            _referenceRemover = referenceRemover;
            _supertypeClearer = supertypeClearer;
        }

        public DeclarationFinder DeclarationFinder => _declarationFinderProvider.DeclarationFinder;

        public void AddModuleToModuleReference(QualifiedModuleName referencingModule, QualifiedModuleName referencedModule)
        {
            _moduleToModuleReferenceManager.AddModuleToModuleReference(referencingModule, referencedModule);
        }

        public void ClearModuleToModuleReferencesFromModule(IEnumerable<QualifiedModuleName> referencingModules)
        {
            _moduleToModuleReferenceManager.ClearModuleToModuleReferencesFromModule(referencingModules);
        }

        public void ClearModuleToModuleReferencesFromModule(QualifiedModuleName referencingModule)
        {
            _moduleToModuleReferenceManager.ClearModuleToModuleReferencesFromModule(referencingModule);
        }

        public void ClearModuleToModuleReferencesToModule(IEnumerable<QualifiedModuleName> referencedModules)
        {
            _moduleToModuleReferenceManager.ClearModuleToModuleReferencesToModule(referencedModules);
        }

        public void ClearModuleToModuleReferencesToModule(QualifiedModuleName referencedModule)
        {
            _moduleToModuleReferenceManager.ClearModuleToModuleReferencesToModule(referencedModule);
        }

        public void ClearSupertypes(IEnumerable<QualifiedModuleName> modules)
        {
            _supertypeClearer.ClearSupertypes(modules);
        }

        public void ClearSupertypes(QualifiedModuleName module)
        {
            _supertypeClearer.ClearSupertypes(module);
        }

        public IReadOnlyCollection<QualifiedModuleName> ModulesReferencedBy(QualifiedModuleName referencingModule)
        {
            return _moduleToModuleReferenceManager.ModulesReferencedBy(referencingModule);
        }

        public IReadOnlyCollection<QualifiedModuleName> ModulesReferencedByAny(IEnumerable<QualifiedModuleName> referencingModules)
        {
            return _moduleToModuleReferenceManager.ModulesReferencedByAny(referencingModules);
        }

        public IReadOnlyCollection<QualifiedModuleName> ModulesReferencing(QualifiedModuleName referencedModule)
        {
            return _moduleToModuleReferenceManager.ModulesReferencing(referencedModule);
        }

        public IReadOnlyCollection<QualifiedModuleName> ModulesReferencingAny(IEnumerable<QualifiedModuleName> referencedModules)
        {
            return _moduleToModuleReferenceManager.ModulesReferencingAny(referencedModules);
        }

        public void RefreshDeclarationFinder()
        {
            _declarationFinderProvider.RefreshDeclarationFinder();
        }

        public void RemoveModuleToModuleReference(QualifiedModuleName referencedModule, QualifiedModuleName referencingModule)
        {
            _moduleToModuleReferenceManager.RemoveModuleToModuleReference(referencedModule, referencingModule);
        }

        public void RemoveReferencesBy(IReadOnlyCollection<QualifiedModuleName> modules, CancellationToken token)
        {
            _referenceRemover.RemoveReferencesBy(modules, token);
        }

        public void RemoveReferencesBy(QualifiedModuleName module, CancellationToken token)
        {
            _referenceRemover.RemoveReferencesBy(module, token);
        }

        public void RemoveReferencesTo(IReadOnlyCollection<QualifiedModuleName> modules, CancellationToken token)
        {
            _referenceRemover.RemoveReferencesTo(modules, token);
        }

        public void RemoveReferencesTo(QualifiedModuleName module, CancellationToken token)
        {
            _referenceRemover.RemoveReferencesTo(module, token);
        }
    }
}
