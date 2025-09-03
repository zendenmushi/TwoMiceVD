using System;
using System.Runtime.InteropServices;

namespace TwoMiceVD.Interop;

internal static class NativeMethods
{
    [DllImport("User32.dll", SetLastError = true)]
    internal static extern bool RegisterRawInputDevices([In] RAWINPUTDEVICE[] pRawInputDevices, uint uiNumDevices, uint cbSize);

    [DllImport("User32.dll", SetLastError = true)]
    internal static extern uint GetRawInputData(IntPtr hRawInput, RawInputCommand uiCommand, IntPtr pData, ref uint pcbSize, uint cbSizeHeader);

    [DllImport("user32.dll")]
    internal static extern bool GetCursorPos(out POINT lpPoint);

    [DllImport("user32.dll")]
    internal static extern bool SetCursorPos(int X, int Y);
}