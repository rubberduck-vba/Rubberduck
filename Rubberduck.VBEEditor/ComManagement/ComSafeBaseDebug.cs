using System;
using System.Collections.Generic;
using System.Diagnostics;
using StreamWriter = System.IO.StreamWriter;
using System.Runtime.InteropServices;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.InternalApi.Common;
using System.IO.Abstractions;

namespace Rubberduck.VBEditor.ComManagement
{
    partial class ComSafeBase
    {
        private struct TraceData
        {
            internal int HashCode { get; set; }
            internal string IUnknownAddress { get; set; }
            internal IEnumerable<string> StackTrace { get; set; }
        }
        private StreamWriter _traceStream;
        private string _traceFilePath;
        private string _directory;
        private readonly object _streamLock = new object();
        private readonly IFileSystem _fileSystem = FileSystemProvider.FileSystem;

        /// <summary>
        /// The first few stack frames come from the ComSafe and thus are not
        /// particularly interesting. Typically, we want to look at the frames
        /// outside the ComSafe. 
        /// </summary>
        private const int StackTraceNumberOfElementsToSkipOnRemoval = 6;
        private const int StackTrackNumberOfElementsToSkipOnAddUpdate = 8;
        private const int StackTraceDepth = 10;

        /// <inheritdoc cref="IComSafe.Serialize"/>
        public void Serialize(string targetDirectory)
        {
            SerializeInternal(targetDirectory);
        }

        [Conditional("TRACE_COM_SAFE")]
        private void SerializeInternal(string targetDirectory)
        {
            lock (_streamLock)
            {
                _directory = targetDirectory;
                var serializeTime = DateTime.UtcNow;
                using (var stream = _fileSystem.File.AppendText(_fileSystem.Path.Combine(_directory,
                    $"COM Safe Content Snapshot {serializeTime:yyyyMMddhhmmss}.csv")))
                {
                    stream.WriteLine(
                        $"Ordinal\tKey\tCOM Wrapper Type\tWrapping Null?\tIUnknown Pointer Address");
                    var i = 0;
                    foreach (var kvp in GetWrappers())
                    {
                        var line = kvp.Value != null
                            ? $"{i++}\t{kvp.Key}\t\"{kvp.Value.GetType().FullName}\"\t\"{kvp.Value.IsWrappingNullReference}\"\t\"{(kvp.Value.IsWrappingNullReference ? "null" : GetPtrAddress(kvp.Value.Target))}\""
                            : $"{i++}\t{kvp.Key}\t\"null\"\t\"null\"\t\"null\"";
                        stream.WriteLine(line);
                    }
                }

                if (_traceStream == null)
                {
                    return;
                }

                _traceStream.Flush();
                _fileSystem.File.Copy(_traceFilePath, _fileSystem.Path.Combine(_directory, $"COM Safe Stack Trace {serializeTime:yyyyMMddhhmmss}.csv"));
            }
        }

        [Conditional("DEBUG")]
        private void TraceDispose()
        {
            lock (_streamLock)
            {
                try
                {
                    if (_traceStream == null)
                    {
                        return;
                    }

                    _traceStream.Close();
                    if (string.IsNullOrWhiteSpace(_directory))
                    {
                        _fileSystem.File.Delete(_traceFilePath);
                    }
                    else
                    {
                        _fileSystem.File.Move(_traceFilePath,
                            _fileSystem.Path.Combine(_directory,
                                _fileSystem.Path.GetFileNameWithoutExtension(_traceFilePath) + " final.csv"));
                    }
                }
                finally
                {
                    _traceStream?.Dispose();
                    _traceStream = null;
                }
            }
        }

        [Conditional("DEBUG")]
        protected void TraceAdd(ISafeComWrapper comWrapper)
        {
            Trace("Add", comWrapper, StackTrackNumberOfElementsToSkipOnAddUpdate);
        }

        [Conditional("DEBUG")]
        protected void TraceUpdate(ISafeComWrapper comWrapper)
        {
            Trace("Update", comWrapper, StackTrackNumberOfElementsToSkipOnAddUpdate);
        }

        [Conditional("DEBUG")]
        protected void TraceRemove(ISafeComWrapper comWrapper, bool wasRemoved)
        {
            var activity = wasRemoved ? "Removed" : "Not removed";
            Trace(activity, comWrapper, StackTraceNumberOfElementsToSkipOnRemoval);
        }

        private readonly object _idLock = new object();
        private int _id;

        [Conditional("DEBUG")]
        private void Trace(string activity, ISafeComWrapper comWrapper, int framesToSkip)
        {
            lock (_streamLock)
            {
                if (_disposed)
                {
                    return;
                }

                if (_traceStream == null)
                {
                    var directory = _fileSystem.Path.GetTempPath();
                    _traceFilePath = _fileSystem.Path.Combine(directory,
                        $"COM Safe Stack Trace {DateTime.UtcNow:yyyyMMddhhmmss}.{GetHashCode()}.csv");
                    _traceStream = _fileSystem.File.AppendText(_traceFilePath);
                    _traceStream.WriteLine(
                        $"Ordinal\tTimestamp\tActivity\tKey\tIUnknown Pointer Address\t{FrameHeaders()}");
                }

                int id;
                lock (_idLock)
                {
                    id = _id++;
                }

                var traceData = new TraceData
                {
                    HashCode = GetComWrapperObjectHashCode(comWrapper),
                    IUnknownAddress = comWrapper.IsWrappingNullReference ? "null" : GetPtrAddress(comWrapper.Target),
                    StackTrace = GetStackTrace(StackTraceDepth, framesToSkip)
                };

                var line =
                    $"{id}\t{DateTime.UtcNow}\t\"{activity}\"\t{traceData.HashCode}\t{traceData.IUnknownAddress}\t\"{string.Join("\"\t\"", traceData.StackTrace)}\"";
                _traceStream.WriteLine(line);
            }
        }

        private static string FrameHeaders()
        {
            var headers = new System.Text.StringBuilder();
            for (var i = 1; i <= StackTraceDepth; i++)
            {
                headers.Append($"Frame {i}\t");
            }

            return headers.ToString();
        }

        protected abstract IDictionary<int, ISafeComWrapper> GetWrappers();

        private static IEnumerable<string> GetStackTrace(int frames, int framesToSkip)
        {
            var list = new List<string>();
            var trace = new StackTrace();
            if (trace.FrameCount < (frames + framesToSkip))
            {
                frames = trace.FrameCount;
            }
            else
            {
                frames += framesToSkip;
            }

            framesToSkip -= 1;
            frames -= 1;

            for (var i = framesToSkip; i < frames; i++)
            {
                var frame = trace.GetFrame(i);
                var type = frame.GetMethod().DeclaringType;

                var typeName = type?.FullName ?? string.Empty;
                var methodName = frame.GetMethod().Name;

                var qualifiedName = $"{typeName}{(typeName.Length > 0 ? "::" : string.Empty)}{methodName}";
                list.Add(qualifiedName);
            }

            return list;
        }

        protected static string GetPtrAddress(object target)
        {
            if (target == null)
            {
                return IntPtr.Zero.ToString();
            }

            if (!Marshal.IsComObject(target))
            {
                return "Not a COM object";
            }

            var pointer = IntPtr.Zero;
            try
            {
                pointer = Marshal.GetIUnknownForObject(target);
            }
            finally
            {
                if (pointer != IntPtr.Zero)
                {
                    Marshal.Release(pointer);
                }
            }

            return pointer.ToString();
        }
    }
}
