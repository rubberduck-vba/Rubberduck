using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;

namespace Rubberduck.UI.WPF
{
    // Taken and adjusted from https://stackoverflow.com/a/5256827/1803692
    /// <summary>
    /// This class extends the capabilities of ObservableCollection.
    /// It adds an additional event, which is raised when the properties of an item stored in the collection change.
    /// </summary>
    public class ObservableViewModelCollection<T> : ObservableCollection<T>
        where T : INotifyPropertyChanged
    {
        public event EventHandler<ElementPropertyChangedEventArgs<T>> ElementPropertyChanged;

        public ObservableViewModelCollection()
        {
            CollectionChanged += UpdateContributorEventHandlers;
        }

        public ObservableViewModelCollection(IEnumerable<T> items) : this()
        {
            foreach (var item in items)
            {
                this.Add(item);
            }
        }

        private void UpdateContributorEventHandlers(object sender, NotifyCollectionChangedEventArgs e)
        {
            foreach (INotifyPropertyChanged added in e.NewItems ?? new List<INotifyPropertyChanged>())
            {
                added.PropertyChanged += OnItemPropertyChanged;
            }
            foreach (INotifyPropertyChanged removed in e.OldItems ?? new List<INotifyPropertyChanged>())
            {
                removed.PropertyChanged -= OnItemPropertyChanged;
            }
        }

        private void OnItemPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            var specializedArgs = new ElementPropertyChangedEventArgs<T>((T)sender, e.PropertyName);
            ElementPropertyChanged?.Invoke(this, specializedArgs);
        }
    }

    public class ElementPropertyChangedEventArgs<T>
    {
        public T Element { get; private set; }
        public string PropertyName { get; private set; }

        public ElementPropertyChangedEventArgs(T element, string propertyName) 
        {
            Element = element;
            PropertyName = propertyName;
        }
    }
}
