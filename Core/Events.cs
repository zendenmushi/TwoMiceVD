using System;
using TwoMiceVD.Interop;

namespace TwoMiceVD.Core;

public class DeviceMovedEventArgs : EventArgs
{
    public string DeviceId { get; }
    public int DeltaX { get; }
    public int DeltaY { get; }
    public RawMouseButtonFlags ButtonFlags { get; }
    public short WheelDelta { get; }
    public short HWheelDelta { get; }

    public DeviceMovedEventArgs(string deviceId, int dx, int dy,
        RawMouseButtonFlags buttonFlags = 0, short wheelDelta = 0, short hWheelDelta = 0)
    {
        DeviceId = deviceId;
        DeltaX = dx;
        DeltaY = dy;
        ButtonFlags = buttonFlags;
        WheelDelta = wheelDelta;
        HWheelDelta = hWheelDelta;
    }
}
