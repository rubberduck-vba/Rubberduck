using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RubberduckTests.Refactoring
{
    public class ConflictDetectionSessionTestsResolver
    {
        public static T Resolve<T>(RubberduckParserState state) where T : class
        {
            return Resolve<T>(state, typeof(T).Name);
        }

        private static T Resolve<T>(RubberduckParserState _state, string name) where T : class
        {
            switch (name)
            {
                case nameof(IConflictDetectionSessionDataFactory):
                    return new ConflictDetectionSessionDataFactory(Resolve<IConflictDetectionDeclarationProxyFactory>(_state)) as T;
                case nameof(IConflictDetectionSessionFactory):
                    return new ConflictDetectionSessionFactory(_state, Resolve<IConflictDetectionSessionDataFactory>(_state), Resolve<IConflictFinderFactory>(_state)) as T;
                case nameof(IConflictDetectionDeclarationProxyFactory):
                    return new ConflictDetectionDeclarationProxyFactory(_state) as T;
                case nameof(IConflictFinderFactory):
                    return new ConflictFinderFactory(_state) as T;
                default:
                    throw new ArgumentException();
            }
        }
    }
}
