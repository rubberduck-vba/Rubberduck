using System;
using System.Runtime.InteropServices;

namespace Rubberduck.SourceControl
{
    //This file was taken from the [Lib2GitSharp project][1] so this enum could be exposed to COM interop.
    //Otherwise, this file remains unchanged.
    //
    //[1]:https://github.com/libgit2/libgit2sharp/
    //
    //The MIT License
    //Copyright (c) LibGit2Sharp contributors
    //
    //Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), 
    //to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, 
    //and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
    //The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
    //
    //THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
    //FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, 
    //WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

    /// <summary>
    /// Calculated status of a filepath in the working directory.
    /// </summary>
    [Flags]
    [ComVisible(true)]
    [Guid("4DDA743E-E3A7-440A-A030-92DF616B2C7B")]
    public enum FileStatus
    {
        /// <summary>
        /// The file doesn't exist.
        /// </summary>
        Nonexistent = (1 << 31),

        /// <summary>
        /// The file hasn't been modified.
        /// </summary>
        Unaltered = 0, /* GIT_STATUS_CURRENT */

        /// <summary>
        /// New file has been added to the Index. It's unknown from the Head.
        /// </summary>
        Added = (1 << 0), /* GIT_STATUS_INDEX_NEW */

        /// <summary>
        /// New version of a file has been added to the Index. A previous version exists in the Head.
        /// </summary>
        Staged = (1 << 1), /* GIT_STATUS_INDEX_MODIFIED */

        /// <summary>
        /// The deletion of a file has been promoted from the working directory to the Index. A previous version exists in the Head.
        /// </summary>
        Removed = (1 << 2), /* GIT_STATUS_INDEX_DELETED */

        /// <summary>
        /// The renaming of a file has been promoted from the working directory to the Index. A previous version exists in the Head.
        /// </summary>
        RenamedInIndex = (1 << 3), /* GIT_STATUS_INDEX_RENAMED */

        /// <summary>
        /// A change in type for a file has been promoted from the working directory to the Index. A previous version exists in the Head.
        /// </summary>
        StagedTypeChange = (1 << 4), /* GIT_STATUS_INDEX_TYPECHANGE */

        /// <summary>
        /// New file in the working directory, unknown from the Index and the Head.
        /// </summary>
        Untracked = (1 << 7), /* GIT_STATUS_WT_NEW */

        /// <summary>
        /// The file has been updated in the working directory. A previous version exists in the Index.
        /// </summary>
        Modified = (1 << 8), /* GIT_STATUS_WT_MODIFIED */

        /// <summary>
        /// The file has been deleted from the working directory. A previous version exists in the Index.
        /// </summary>
        Missing = (1 << 9), /* GIT_STATUS_WT_DELETED */

        /// <summary>
        /// The file type has been changed in the working directory. A previous version exists in the Index.
        /// </summary>
        TypeChanged = (1 << 10), /* GIT_STATUS_WT_TYPECHANGE */

        /// <summary>
        /// The file has been renamed in the working directory.  The previous version at the previous name exists in the Index.
        /// </summary>
        RenamedInWorkDir = (1 << 11), /* GIT_STATUS_WT_RENAMED */

        /// <summary>
        /// The file is unreadable in the working directory.
        /// </summary>
        Unreadable = (1 << 12), /* GIT_STATUS_WT_UNREADABLE */

        /// <summary>
        /// The file is <see cref="Untracked"/> but its name and/or path matches an exclude pattern in a <c>gitignore</c> file.
        /// </summary>
        Ignored = (1 << 14), /* GIT_STATUS_IGNORED */
    }
}
