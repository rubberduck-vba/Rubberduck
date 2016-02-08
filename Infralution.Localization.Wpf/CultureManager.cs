//
//      FILE:   CultureManager.cs.
//
// COPYRIGHT:   Copyright 2008 
//              Infralution
//
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Threading;
using System.Windows;
using System.Windows.Markup;
using System.Windows.Forms;
using System.Reflection;
using Infralution.Localization.Wpf.Properties;
using System.Runtime.InteropServices;
namespace Infralution.Localization.Wpf
{

    /// <summary>
    /// Provides the ability to change the UICulture for WPF Windows and controls
    /// dynamically.  
    /// </summary>
    /// <remarks>
    /// XAML elements that use the <see cref="ResxExtension"/> are automatically
    /// updated when the <see cref="CultureManager.UICulture"/> property is changed.
    /// </remarks>
    public static class CultureManager
    {
        #region Static Member Variables

        /// <summary>
        /// Current UICulture of the application
        /// </summary>
        private static CultureInfo _uiCulture;

        /// <summary>
        /// The active design time culture selection window (if any)
        /// </summary>
        private static CultureSelectWindow _cultureSelectWindow;

        /// <summary>
        /// The active task bar notify icon for design time culture selection (if any)
        /// </summary>
        private static NotifyIcon _notifyIcon;

        /// <summary>
        /// The window handle for the notify icon
        /// </summary>
        private static IntPtr _notifyIconHandle;

        /// <summary>
        /// Should the <see cref="Thread.CurrentCulture"/> be changed when the
        /// <see cref="UICulture"/> changes.
        /// </summary>
        private static bool _synchronizeThreadCulture = true;

        #endregion

        #region Public Interface

        /// <summary>
        /// Raised when the <see cref="UICulture"/> is changed
        /// </summary>
        /// <remarks>
        /// Since this event is static if the client object does not detach from the event a reference
        /// will be maintained to the client object preventing it from being garbage collected - thus
        /// causing a potential memory leak. 
        /// </remarks>
        public static event EventHandler UICultureChanged;

        /// <summary>
        /// Sets the UICulture for the WPF application and raises the <see cref="UICultureChanged"/>
        /// event causing any XAML elements using the <see cref="ResxExtension"/> to automatically
        /// update
        /// </summary>
        public static CultureInfo UICulture
        {
            get
            {
                if (_uiCulture == null)
                {
                    _uiCulture = Thread.CurrentThread.CurrentUICulture;
                }
                return _uiCulture;
            }
            set
            {
                if (value != UICulture)
                {
                    _uiCulture = value;
                    Thread.CurrentThread.CurrentUICulture = value;
                    if (SynchronizeThreadCulture)
                    {
                        SetThreadCulture(value);
                    }
                    UICultureExtension.UpdateAllTargets();
                    ResxExtension.UpdateAllTargets();
                    if (UICultureChanged != null)
                    {
                        UICultureChanged(null, EventArgs.Empty);
                    }
                }
            }
        }

        /// <summary>
        /// If set to true then the <see cref="Thread.CurrentCulture"/> property is changed
        /// to match the current <see cref="UICulture"/>
        /// </summary>
        public static bool SynchronizeThreadCulture
        {
            get { return _synchronizeThreadCulture; }
            set
            {
                _synchronizeThreadCulture = value;
                if (value)
                {
                    SetThreadCulture(UICulture);
                }
            }
        }

        #endregion

        #region Local Methods

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        private class NOTIFYICONDATA
        {
            public int cbSize = Marshal.SizeOf(typeof(NOTIFYICONDATA));
            public IntPtr hWnd;
            public int uID;
            public int uFlags;
            public int uCallbackMessage;
            public IntPtr hIcon;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)]
            public string szTip;
            public int dwState;
            public int dwStateMask;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)]
            public string szInfo;
            public int uTimeoutOrVersion;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 64)]
            public string szInfoTitle;
            public int dwInfoFlags;
        }

        [DllImport("shell32.dll", CharSet = CharSet.Auto)]
        private static extern int Shell_NotifyIcon(int message, NOTIFYICONDATA pnid);

        /// <summary>
        /// Set the thread culture to the given culture
        /// </summary>
        /// <param name="value">The culture to set</param>
        /// <remarks>If the culture is neutral then creates a specific culture</remarks>
        private static void SetThreadCulture(CultureInfo value)
        {
            if (value.IsNeutralCulture)
            {
                Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture(value.Name);
            }
            else
            {
                Thread.CurrentThread.CurrentCulture = value;
            }
        }

        /// <summary>
        /// Show the UICultureSelector to allow selection of the active UI culture
        /// </summary>
        internal static void ShowCultureNotifyIcon()
        {
            if (_notifyIcon == null)
            {
                ToolStripMenuItem menuItem;

                _notifyIcon = new NotifyIcon();
                _notifyIcon.Icon = Resources.UICultureIcon;
                _notifyIcon.MouseClick += new MouseEventHandler(OnCultureNotifyIconMouseClick);
                _notifyIcon.MouseDoubleClick += new MouseEventHandler(OnCultureNotifyIconMouseDoubleClick);
                _notifyIcon.Text = Resources.UICultureSelectText;
                ContextMenuStrip menuStrip = new ContextMenuStrip();

                // separator
                //
                menuStrip.Items.Add(new ToolStripSeparator());

                // add menu to open culture select window
                //
                menuItem = new ToolStripMenuItem(Resources.OtherCulturesMenu);
                menuItem.Click += new EventHandler(OnCultureSelectMenuClick);
                menuStrip.Items.Add(menuItem);

                // add menu to exit the designer for VS2012/2013
                //
                if (AppDomain.CurrentDomain.FriendlyName == "XDesProc.exe")
                {
                    menuItem = new ToolStripMenuItem(Resources.ExitDesignerMenu);
                    menuItem.Click += new EventHandler(OnExitDesignerMenuClick);
                    menuStrip.Items.Add(menuItem);
                }

                menuStrip.Opening += OnMenuStripOpening;
                _notifyIcon.ContextMenuStrip = menuStrip;
                _notifyIcon.Visible = true;

                // Save the window handle associated with the notify icon - note that the window
                // is destroyed before the ProcessExit event gets called so calling NotifyIcon.Dispose
                // within the ProcessExit event handler doesn't work because the window handle has been
                // set to zero by that stage
                //
                FieldInfo fieldInfo = typeof(NotifyIcon).GetField("window", BindingFlags.Instance | BindingFlags.NonPublic);
                if (fieldInfo != null)
                {
                    NativeWindow iconWindow = fieldInfo.GetValue(_notifyIcon) as NativeWindow;
                    _notifyIconHandle = iconWindow.Handle;
                }

                AppDomain.CurrentDomain.ProcessExit += OnDesignerExit;
            }
        }

        /// <summary>
        /// Remove the culture notify icon when the designer process exits
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private static void OnDesignerExit(object sender, EventArgs e)
        {
            // By the time the ProcessExit event is called the window associated with the
            // notify icon has been destroyed - and a bug in the NotifyIcon class means that
            // the notify icon is not removed. This works around the issue by saving the 
            // window handle when the NotifyIcon is created and then calling the Shell_NotifyIcon
            // method ourselves to remove the icon from the tray
            //
            if (_notifyIconHandle != IntPtr.Zero)
            {
                NOTIFYICONDATA iconData = new NOTIFYICONDATA();
                iconData.uCallbackMessage = 2048;
                iconData.uFlags = 1;
                iconData.hWnd = _notifyIconHandle;
                iconData.uID = 1;
                iconData.hIcon = IntPtr.Zero;
                iconData.szTip = null;
                Shell_NotifyIcon(2, iconData);
            }
        }

        /// <summary>
        /// Display the CultureSelectWindow to allow the user to select the UICulture
        /// </summary>
        private static void DisplayCultureSelectWindow()
        {
            if (_cultureSelectWindow == null)
            {
                _cultureSelectWindow = new CultureSelectWindow();
                _cultureSelectWindow.Title = _notifyIcon.Text;
                _cultureSelectWindow.Closed += new EventHandler(OnCultureSelectWindowClosed);
                _cultureSelectWindow.Show();
            }
        }

        /// <summary>
        /// Is there already an entry for the culture in the context menu
        /// </summary>
        /// <param name="culture">The culture to check</param>
        /// <returns>True if there is a menu</returns>
        private static bool CultureMenuExists(CultureInfo culture)
        {
            foreach (ToolStripItem item in _notifyIcon.ContextMenuStrip.Items)
            {
                CultureInfo itemCulture = item.Tag as CultureInfo;
                if (itemCulture != null && itemCulture.Name == culture.Name)
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Add a menu item to the NotifyIcon for the current UICulture
        /// </summary>
        /// <param name="culture"></param>
        private static void AddCultureMenuItem(CultureInfo culture)
        {
            if (!CultureMenuExists(culture))
            {
                ContextMenuStrip menuStrip = _notifyIcon.ContextMenuStrip;
                ToolStripMenuItem menuItem = new ToolStripMenuItem(culture.DisplayName);
                menuItem.Checked = true;
                menuItem.CheckOnClick = true;
                menuItem.Tag = culture;
                menuItem.CheckedChanged += new EventHandler(OnCultureMenuCheckChanged);
                menuStrip.Items.Insert(0, menuItem);
            }
        }

        /// <summary>
        /// Update the notify icon menu
        /// </summary>
        private static void OnMenuStripOpening(object sender, System.ComponentModel.CancelEventArgs e)
        {

            // ensure the current culture is always on the menu
            //
            AddCultureMenuItem(UICulture);

            // Add the design time cultures
            //
            List<CultureInfo> designTimeCultures = ResxExtension.GetDesignTimeCultures();
            foreach (CultureInfo culture in designTimeCultures)
            {
                AddCultureMenuItem(culture);
            }

            ContextMenuStrip menuStrip = _notifyIcon.ContextMenuStrip;
            foreach (ToolStripItem item in menuStrip.Items)
            {
                ToolStripMenuItem menuItem = item as ToolStripMenuItem;
                if (menuItem != null)
                {
                    menuItem.Checked = (menuItem.Tag == UICulture);
                }
            }
        }

        /// <summary>
        /// Display the context menu for left clicks (right clicks are handled automatically)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private static void OnCultureNotifyIconMouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                MethodInfo methodInfo = typeof(NotifyIcon).GetMethod("ShowContextMenu",
                         BindingFlags.Instance | BindingFlags.NonPublic);
                methodInfo.Invoke(_notifyIcon, null);
            }
        }

        /// <summary>
        /// Display the CultureSelectWindow when the user double clicks on the NotifyIcon
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private static void OnCultureNotifyIconMouseDoubleClick(object sender, MouseEventArgs e)
        {
            DisplayCultureSelectWindow();
        }

        /// <summary>
        /// Display the CultureSelectWindow when the user selects the menu option
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private static void OnCultureSelectMenuClick(object sender, EventArgs e)
        {
            DisplayCultureSelectWindow();
        }

        /// <summary>
        /// Exit the designer process
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private static void OnExitDesignerMenuClick(object sender, EventArgs e)
        {
            System.Windows.Application.Current.Shutdown();
        }

        /// <summary>
        /// Handle change of culture via the NotifyIcon menu
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private static void OnCultureMenuCheckChanged(object sender, EventArgs e)
        {
            ToolStripMenuItem menuItem = sender as ToolStripMenuItem;
            if (menuItem.Checked)
            {
                UICulture = menuItem.Tag as CultureInfo;
            }
        }

        /// <summary>
        /// Handle close of the culture select window
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private static void OnCultureSelectWindowClosed(object sender, EventArgs e)
        {
            _cultureSelectWindow = null;
        }

        #endregion

    }

}
