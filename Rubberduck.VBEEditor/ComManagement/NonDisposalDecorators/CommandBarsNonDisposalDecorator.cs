using System.Collections;
using System.Collections.Generic;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.ComManagement.NonDisposalDecorators
{
    public class CommandBarsNonDisposalDecorator<T> : NonDisposalDecoratorBase<T>, ICommandBars
        where T : ICommandBars
    {
        public CommandBarsNonDisposalDecorator(T commandBars)
            : base(commandBars)
        { }

        public IEnumerator<ICommandBar> GetEnumerator()
        {
            return WrappedItem.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return ((IEnumerable) WrappedItem).GetEnumerator();
        }

        public int Count => WrappedItem.Count;

        public ICommandBar this[object index] => WrappedItem[index];

        public ICommandBar Add(string name)
        {
            return WrappedItem.Add(name);
        }

        public ICommandBar Add(string name, CommandBarPosition position)
        {
            return WrappedItem.Add(name, position);
        }

        public ICommandBarControl FindControl(int id)
        {
            return WrappedItem.FindControl(id);
        }

        public ICommandBarControl FindControl(ControlType type, int id)
        {
            return WrappedItem.FindControl(type, id);
        }
    }
}