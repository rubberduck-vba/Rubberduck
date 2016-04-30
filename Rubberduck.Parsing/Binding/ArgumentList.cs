using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Parsing.Binding
{
    public sealed class ArgumentList
    {
        private readonly List<ArgumentListArgument> _arguments;

        public ArgumentList()
        {
            _arguments = new List<ArgumentListArgument>();
        }

        public void AddArgument(ArgumentListArgumentType argumentType)
        {
            _arguments.Add(new ArgumentListArgument(argumentType));
        }

        public bool HasArguments
        {
            get
            {
                return HasRequiredPositionalArgument|| HasNamedArguments;
            }
        }

        public bool HasNamedArguments
        {
            get
            {
                return _arguments.Any(a => a.ArgumentType == ArgumentListArgumentType.Named);
            }
        }

        public bool HasRequiredPositionalArgument
        {
            get
            {
                return _arguments.Any(a => a.ArgumentType == ArgumentListArgumentType.Positional);
            }
        }
    }
}
