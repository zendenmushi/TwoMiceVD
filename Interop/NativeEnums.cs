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

// Button and wheel flags for RAWMOUSE.usButtonFlags
[Flags]
public enum RawMouseButtonFlags : ushort
{
    RI_MOUSE_LEFT_BUTTON_DOWN   = 0x0001,
    RI_MOUSE_LEFT_BUTTON_UP     = 0x0002,
    RI_MOUSE_RIGHT_BUTTON_DOWN  = 0x0004,
    RI_MOUSE_RIGHT_BUTTON_UP    = 0x0008,
    RI_MOUSE_MIDDLE_BUTTON_DOWN = 0x0010,
    RI_MOUSE_MIDDLE_BUTTON_UP   = 0x0020,
    RI_MOUSE_BUTTON_4_DOWN      = 0x0040,
    RI_MOUSE_BUTTON_4_UP        = 0x0080,
    RI_MOUSE_BUTTON_5_DOWN      = 0x0100,
    RI_MOUSE_BUTTON_5_UP        = 0x0200,
    RI_MOUSE_WHEEL              = 0x0400,
    RI_MOUSE_HWHEEL             = 0x0800,
}

// GetRawInputDeviceInfo で使用するコマンド
internal enum RawInputDeviceInfoCommand : uint
{
    RIDI_DEVICENAME = 0x20000007,
}
