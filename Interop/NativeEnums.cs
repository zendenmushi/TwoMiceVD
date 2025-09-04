using System;

namespace TwoMiceVD.Interop;

internal enum RawInputDeviceFlags : uint
{
    INPUTSINK = 0x00000100,
}

internal enum RawInputType : uint
{
    RIM_TYPEMOUSE = 0,
    RIM_TYPEKEYBOARD = 1,
    RIM_TYPEHID = 2
}

internal enum RawInputCommand : uint
{
    RID_INPUT = 0x10000003,
    RID_HEADER = 0x10000005
}

[Flags]
internal enum RawMouseFlags : ushort
{
    MOUSE_MOVE_RELATIVE = 0x00,
    MOUSE_MOVE_ABSOLUTE = 0x01,
    MOUSE_VIRTUAL_DESKTOP = 0x02,
    MOUSE_ATTRIBUTES_CHANGED = 0x04
}

// GetRawInputDeviceInfo で使用するコマンド
internal enum RawInputDeviceInfoCommand : uint
{
    RIDI_DEVICENAME = 0x20000007,
}
