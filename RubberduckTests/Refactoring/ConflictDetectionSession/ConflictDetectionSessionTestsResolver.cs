using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using System;

namespace RubberduckTests.Refactoring
{
    public class ConflictDetectorTestsResolver
    {
        public static T Resolve<T>(RubberduckParserState state) where T : class
        {
            return Resolve<T>(state, typeof(T).Name);
        }

        private static T Resolve<T>(RubberduckParserState _state, string name) where T : class
        {
            switch (name)
            {
                case nameof(IDeclarationProxyFactory):
                    return new ConflictDetectionDeclarationProxyFactory(_state) as T;
                case nameof(IConflictFinderFactory):
                    return new ConflictFinderFactory(_state, 
                                                        Resolve<IDeclarationProxyFactory>(_state)) as T;
                case nameof(IConflictDetectorFactory):
                    return new ConflictDetectorFactory(_state,
                                                        Resolve<IConflictFinderFactory>(_state),
                                                        Resolve<IDeclarationProxyFactory>(_state)) as T;
                case nameof(IConflictSessionFactory):
                    return new ConflictSessionFactory(_state, Resolve<IDeclarationProxyFactory>(_state), Resolve<IConflictDetectorFactory>(_state)) as T;
                default:
                    throw new ArgumentException();
            }
        }
    }
}
