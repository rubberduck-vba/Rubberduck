namespace Rubberduck.Common.WinAPI
{
    enum DeviceNotification
    {
        DEVICE_NOTIFY_WINDOW_HANDLE = 0x00000000,           // The hRecipient parameter is a window handle
        DEVICE_NOTIFY_SERVICE_HANDLE = 0x00000001,          // The hRecipient parameter is a service status handle
        DEVICE_NOTIFY_ALL_INTERFACE_CLASSES = 0x00000004    // Notifies the recipient of device interface events for all device interface classes. (The dbcc_classguid member is ignored.)
                                                            // This value can be used only if the dbch_devicetype member is DBT_DEVTYP_DEVICEINTERFACE.
    }
}
