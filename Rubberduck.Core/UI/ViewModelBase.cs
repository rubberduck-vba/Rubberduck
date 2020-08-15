using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Media.Imaging;
using Rubberduck.Properties;

namespace Rubberduck.UI
{
    public abstract class ViewModelBase : INotifyPropertyChanged, INotifyDataErrorInfo
    {
        public event PropertyChangedEventHandler PropertyChanged;
        public event EventHandler<DataErrorsChangedEventArgs> ErrorsChanged;

        private readonly IDictionary<string, List<string>> _errors = new ConcurrentDictionary<string, List<string>>();

        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        protected static BitmapImage GetImageSource(Bitmap image)
        {
            using (var memory = new MemoryStream())
            {
                image.Save(memory, ImageFormat.Png);
                memory.Position = 0;
                var bitmapImage = new BitmapImage();
                bitmapImage.BeginInit();
                bitmapImage.StreamSource = memory;
                bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                bitmapImage.EndInit();
                bitmapImage.Freeze();

                return bitmapImage;
            }
        }

        protected virtual void OnErrorsChanged(string propertyName = null)
        {
            ErrorsChanged?.Invoke(this, new DataErrorsChangedEventArgs(propertyName));
            OnPropertyChanged(nameof(HasErrors));
        }

        public IEnumerable GetErrors(string propertyName)
        {
            if (propertyName != null)
            {
                return _errors.TryGetValue(propertyName, out var errorList)
                    ? errorList
                    : null;
            }
            return null;
        }

        public bool HasErrors => _errors.Any();

        /// <summary>
        /// Replaces all errors for a property and notifies all consumers.
        /// </summary>
        /// <param name="propertyName">Name of the property</param>
        /// <param name="errorTexts">List of texts describing each error</param>
        protected void SetErrors(string propertyName, List<string> errorTexts)
        {
            if (propertyName == null)
            {
                return;
            }

            _errors[propertyName] = errorTexts;
            OnErrorsChanged(propertyName);
        }

        /// <summary>
        /// Adds a single error for a property and notifies all consumers.
        /// </summary>
        /// <param name="propertyName">Name of the property</param>
        /// <param name="errorText">Text describing the error</param>
        protected void AddError(string propertyName, string errorText)
        {
            if (propertyName == null)
            {
                return;
            }

            if (_errors.TryGetValue(propertyName, out var errorList))
            {
                errorList.Add(errorText);
            }
            else
            {
                _errors.Add(propertyName, new List<string>{errorText});
            }

            OnErrorsChanged(propertyName);
        }

        /// <summary>
        /// Clears all errors for a property and notifies all consumers.
        /// <remarks>
        /// If no argument or null is provided, all errors will be cleared.
        /// </remarks>
        /// </summary>
        /// <param name="propertyName">Name of the property</param>
        protected void ClearErrors(string propertyName = null)
        {
            if (!_errors.Any())
            {
                return;
            }

            if (propertyName == null)
            {
                var errorProperties = _errors.Keys.ToList();
                _errors.Clear();
                foreach (var errorPropertyName in errorProperties)
                {
                    OnErrorsChanged(errorPropertyName);
                }
            }
            else if (_errors.ContainsKey(propertyName))
            {
                _errors.Remove(propertyName);
                OnErrorsChanged(propertyName);
            }
        }
    }
}
