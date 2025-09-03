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

[StructLayout(LayoutKind.Explicit)]
internal struct RAWMOUSE
{
    [FieldOffset(0)]
    public MOUSE mouse;
}

[StructLayout(LayoutKind.Sequential)]
internal struct MOUSE
{
    public RawMouseFlags usFlags;
    public uint ulButtons;
    public ushort usButtonFlags;
    public ushort usButtonData;
    public uint ulRawButtons;
    public int lLastX;
    public int lLastY;
    public uint ulExtraInformation;
}

[StructLayout(LayoutKind.Sequential)]
internal struct POINT
{
    public int X;
    public int Y;
}