using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Common;
using System;

namespace Rubberduck.Refactorings
{
    public interface IConflictFinderFactory
    {
        IConflictFinder Create(DeclarationType declarationType);
    }

    public class ConflictFinderFactory : IConflictFinderFactory
    {
        private readonly ConflictFinderProject _conflictFinderProject;
        private readonly ConflictFinderModule _conflictFinderModule;
        private readonly ConflictFinderMembers _conflictFinderMembers;
        private readonly ConflictFinderProperties _conflictFinderProperties;
        private readonly ConflictFinderNonMembers _conflictFinderNonMembers;
        private readonly ConflictFinderEvent _conflictFinderEvent;
        private readonly ConflictFinderParameter _conflictFinderParameter;
        private readonly ConflictFinderUDT _conflictFinderUDT;
        private readonly ConflictFinderEnum _conflictFinderEnum;

        public ConflictFinderFactory(IDeclarationFinderProvider declarationFinderProvider)
        {
            _conflictFinderProject = new ConflictFinderProject(declarationFinderProvider);
            _conflictFinderModule = new ConflictFinderModule(declarationFinderProvider);
            _conflictFinderMembers = new ConflictFinderMembers(declarationFinderProvider);
            _conflictFinderProperties = new ConflictFinderProperties(declarationFinderProvider);
            _conflictFinderNonMembers = new ConflictFinderNonMembers(declarationFinderProvider);
            _conflictFinderEvent = new ConflictFinderEvent(declarationFinderProvider);
            _conflictFinderParameter = new ConflictFinderParameter(declarationFinderProvider);
            _conflictFinderUDT = new ConflictFinderUDT(declarationFinderProvider);
            _conflictFinderEnum = new ConflictFinderEnum(declarationFinderProvider);
        }

        public IConflictFinder Create(DeclarationType declarationType)
        {
            switch (declarationType)
            {
                case DeclarationType.Project:
                    return _conflictFinderProject;
                case DeclarationType.Module:
                case DeclarationType.ProceduralModule:
                case DeclarationType.ClassModule:
                    return _conflictFinderModule;
                case DeclarationType.Function:
                case DeclarationType.Procedure:
                    return _conflictFinderMembers;
                case DeclarationType.Property:
                case DeclarationType.PropertyGet:
                case DeclarationType.PropertySet:
                case DeclarationType.PropertyLet:
                    return _conflictFinderProperties;
                case DeclarationType.Variable:
                case DeclarationType.Constant:
                    return _conflictFinderNonMembers;
                case DeclarationType.Event:
                    return _conflictFinderEvent;
                case DeclarationType.Parameter:
                    return _conflictFinderParameter;
                case DeclarationType.UserDefinedType:
                case DeclarationType.UserDefinedTypeMember:
                    return _conflictFinderUDT;
                case DeclarationType.Enumeration:
                case DeclarationType.EnumerationMember:
                    return _conflictFinderEnum;
                default:
                    throw new ArgumentException();
            }
        }
    }
}
