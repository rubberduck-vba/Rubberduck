using System;
using System.Collections;
using System.Collections.Generic;

namespace Rubberduck.VBEditor.ComManagement.TypeLibs.Utility
{
    /// <summary>
    /// A disposable list that encapsulates the disposing of its elements
    /// </summary>
    /// <typeparam name="T"></typeparam>
    internal class DisposableList<T> : IList<T>, IDisposable
        where T : IDisposable
    {
        private readonly IList<T> _list = new List<T>();

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        ~DisposableList()
        {
            System.Diagnostics.Debug.Print("DisposableList:: Finalize called");
        }

        private bool _isDisposed;
        protected virtual void Dispose(bool disposing)
        {
            if (_isDisposed || !disposing)
            {
                return;
            }
            _isDisposed = true;

            foreach (var element in _list)
            {
                element.Dispose();
            }
        }

        public IEnumerator<T> GetEnumerator() => _list.GetEnumerator();
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

        public void Add(T item) => _list.Add(item);
        public void Clear() => _list.Clear();
        public bool Contains(T item) => _list.Contains(item);
        public void CopyTo(T[] array, int arrayIndex) => _list.CopyTo(array, arrayIndex);
        public bool Remove(T item) => _list.Remove(item);
        public int Count => _list.Count;
        public bool IsReadOnly => _list.IsReadOnly;

        public int IndexOf(T item) => _list.IndexOf(item);
        public void Insert(int index, T item) => _list.Insert(index, item);
        public void RemoveAt(int index) => _list.RemoveAt(index);
        public T this[int index]
        {
            get => _list[index];
            set => _list[index] = value;
        }
    }
}
