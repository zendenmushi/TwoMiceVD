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

    [DllImport("user32.dll")]
    internal static extern IntPtr GetForegroundWindow();

    // 端末名（デバイスパス）取得用
    internal const uint RIDI_DEVICENAME = 0x20000007;

    [DllImport("User32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    internal static extern uint GetRawInputDeviceInfo(IntPtr hDevice, uint uiCommand, IntPtr pData, ref uint pcbSize);

    // Window style helpers
    internal const int GWL_EXSTYLE = -20;
    internal const int SWP_NOSIZE = 0x0001;
    internal const int SWP_NOMOVE = 0x0002;
    internal const int SWP_NOZORDER = 0x0004;
    internal const int SWP_FRAMECHANGED = 0x0020;

    [DllImport("user32.dll", SetLastError = true)]
    internal static extern int GetWindowLong(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll", SetLastError = true)]
    internal static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

    [DllImport("user32.dll", SetLastError = true)]
    internal static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
}
