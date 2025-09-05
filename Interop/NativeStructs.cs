using System;
using System.Runtime.InteropServices;

namespace TwoMiceVD.Interop;

[StructLayout(LayoutKind.Sequential)]
internal struct RAWINPUTDEVICE
{
    public ushort usUsagePage;
    public ushort usUsage;
    public RawInputDeviceFlags dwFlags;
    public IntPtr hwndTarget;
}

[StructLayout(LayoutKind.Sequential)]
internal struct RAWINPUTHEADER
{
    public RawInputType dwType;
    public uint dwSize;
    public IntPtr hDevice;
    public IntPtr wParam;
}

[StructLayout(LayoutKind.Sequential)]
internal struct RAWINPUT
{
    public RAWINPUTHEADER header;
    public RAWMOUSE data;
}

// Correct RAWMOUSE layout with explicit union for button fields
[StructLayout(LayoutKind.Sequential)]
internal struct RAWMOUSE
{
    public RawMouseFlags usFlags;
    public RAWMOUSE_BUTTONS Buttons; // union of ulButtons / (usButtonFlags, usButtonData)
    public uint ulRawButtons;
    public int lLastX;
    public int lLastY;
    public uint ulExtraInformation;
}

[StructLayout(LayoutKind.Explicit)]
internal struct RAWMOUSE_BUTTONS
{
    [FieldOffset(0)] public uint ulButtons;
    [FieldOffset(0)] public ushort usButtonFlags;
    [FieldOffset(2)] public ushort usButtonData;
}

[StructLayout(LayoutKind.Sequential)]
internal struct POINT
{
    public int X;
    public int Y;
}
