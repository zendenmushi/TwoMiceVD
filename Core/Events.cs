using System;

namespace TwoMiceVD.Core;

public class DeviceMovedEventArgs : EventArgs
{
    public string DeviceId { get; }
    public int DeltaX { get; }
    public int DeltaY { get; }

    public DeviceMovedEventArgs(string deviceId, int dx, int dy)
    {
        DeviceId = deviceId;
        DeltaX = dx;
        DeltaY = dy;
    }
}