using System;
using System.ComponentModel;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using TwoMiceVD.Core;
using TwoMiceVD.Interop;
using static TwoMiceVD.Interop.NativeMethods;

namespace TwoMiceVD.Input;

public class RawInputManager : IDisposable
{
    private readonly HiddenForm? _hiddenForm;

    public event EventHandler<DeviceMovedEventArgs>? DeviceMoved;
    public event EventHandler<DeviceConnectionChangedEventArgs>? DeviceConnectionChanged;

    public RawInputManager()
    {
        _hiddenForm = new HiddenForm();
        _hiddenForm.RawInputReceived += OnRawInputReceived;
        _hiddenForm.DeviceChangeReceived += OnRawInputDeviceChangeReceived;
        
        // フォームのハンドル作成を明示的に実行
        _hiddenForm.Initialize();
        
        // ハンドルが作成されてからRaw Input を登録
        if (!_hiddenForm.IsHandleCreated)
        {
            throw new InvalidOperationException("HiddenFormのハンドル作成に失敗しました");
        }
        
        RegisterRawInput();
    }

    /// <summary>
    /// 現在接続中のマウスデバイス（Raw Input）のデバイスパス一覧を取得
    /// </summary>
    public string[] GetCurrentlyConnectedMouseDevicePaths()
    {
        try
        {
            uint count = 0;
            uint cb = (uint)Marshal.SizeOf<RAWINPUTDEVICELIST>();
            // 1st call to get count
            GetRawInputDeviceList(IntPtr.Zero, ref count, cb);
            if (count == 0) return Array.Empty<string>();

            IntPtr pList = Marshal.AllocHGlobal((int)(count * cb));
            try
            {
                uint ret = GetRawInputDeviceList(pList, ref count, cb);
                if (ret == uint.MaxValue || ret == 0) return Array.Empty<string>();

                var result = new System.Collections.Generic.List<string>();
                for (int i = 0; i < ret; i++)
                {
                    IntPtr pItem = IntPtr.Add(pList, i * (int)cb);
                    var item = Marshal.PtrToStructure<RAWINPUTDEVICELIST>(pItem);
                    if (item.dwType == RawInputType.RIM_TYPEMOUSE)
                    {
                        string id = GetDeviceId(item.hDevice);
                        if (!string.IsNullOrWhiteSpace(id))
                            result.Add(id);
                    }
                }
                return result.ToArray();
            }
            finally
            {
                Marshal.FreeHGlobal(pList);
            }
        }
        catch
        {
            return Array.Empty<string>();
        }
    }

    private void RegisterRawInput()
    {
        RAWINPUTDEVICE[] rid = new RAWINPUTDEVICE[1];
        rid[0].usUsagePage = 0x01;
        rid[0].usUsage = 0x02; // マウス
        rid[0].dwFlags = RawInputDeviceFlags.INPUTSINK | RawInputDeviceFlags.DEVNOTIFY;
        rid[0].hwndTarget = _hiddenForm!.Handle;
        if (!RegisterRawInputDevices(rid, (uint)rid.Length, (uint)Marshal.SizeOf<RAWINPUTDEVICE>()))
        {
            throw new Win32Exception(Marshal.GetLastWin32Error());
        }
    }

    private void OnRawInputReceived(object? sender, RawInputEventArgs e)
    {
        ProcessRawInput(e.LParam);
    }

    private void ProcessRawInput(IntPtr lParam)
    {
        uint size = 0;
        GetRawInputData(lParam, RawInputCommand.RID_INPUT, IntPtr.Zero, ref size, (uint)Marshal.SizeOf<RAWINPUTHEADER>());

        IntPtr buffer = Marshal.AllocHGlobal((int)size);
        try
        {
            if (GetRawInputData(lParam, RawInputCommand.RID_INPUT, buffer, ref size, (uint)Marshal.SizeOf<RAWINPUTHEADER>()) == size)
            {
                RAWINPUT rawInput = Marshal.PtrToStructure<RAWINPUT>(buffer);
                if (rawInput.header.dwType == RawInputType.RIM_TYPEMOUSE)
                {
                    string deviceId = GetDeviceId(rawInput.header.hDevice);
                    int deltaX = rawInput.data.lLastX;
                    int deltaY = rawInput.data.lLastY;

                    // Button/wheel flags (from union)
                    var btnFlags = (RawMouseButtonFlags)rawInput.data.Buttons.usButtonFlags;

                    // Extract wheel deltas when present (signed short in usButtonData)
                    short wheelDelta = 0;
                    short hWheelDelta = 0;
                    if ((btnFlags & RawMouseButtonFlags.RI_MOUSE_WHEEL) != 0)
                        wheelDelta = unchecked((short)rawInput.data.Buttons.usButtonData);
                    if ((btnFlags & RawMouseButtonFlags.RI_MOUSE_HWHEEL) != 0)
                        hWheelDelta = unchecked((short)rawInput.data.Buttons.usButtonData);

                    // Treat button clicks and wheel rotations as activity (even without movement)
                    bool hasButtonOrWheel = btnFlags != 0;

                    if (deltaX != 0 || deltaY != 0 || hasButtonOrWheel)
                    {
                        DeviceMoved?.Invoke(this,
                            new DeviceMovedEventArgs(deviceId, deltaX, deltaY, btnFlags, wheelDelta, hWheelDelta));
                    }
                }
            }
        }
        finally
        {
            Marshal.FreeHGlobal(buffer);
        }
    }

    private string GetDeviceId(IntPtr hDevice)
    {
        // 安定したデバイス識別子として、Raw Input のデバイス名（パス）を取得
        try
        {
            uint size = 0;
            // まず必要サイズを取得（戻り値は文字数）
            GetRawInputDeviceInfo(hDevice, (uint)RawInputDeviceInfoCommand.RIDI_DEVICENAME, IntPtr.Zero, ref size);
            if (size > 0)
            {
                // Unicode 文字数をバイト数に換算して領域確保
                IntPtr buffer = Marshal.AllocHGlobal((int)size * 2);
                try
                {
                    uint result = GetRawInputDeviceInfo(hDevice, (uint)RawInputDeviceInfoCommand.RIDI_DEVICENAME, buffer, ref size);
                    if (result > 0)
                    {
                        string? name = Marshal.PtrToStringUni(buffer);
                        if (!string.IsNullOrEmpty(name))
                        {
                            return name;
                        }
                    }
                }
                finally
                {
                    Marshal.FreeHGlobal(buffer);
                }
            }
        }
        catch { /* 失敗時はフォールバック */ }

        // フォールバック：ハンドル由来（セッション間で不安定）
        return $"Device_{hDevice.ToInt64():X}";
    }

    public void Dispose()
    {
        _hiddenForm?.Dispose();
    }

    /// <summary>
    /// WM_INPUT メッセージを受信するための隠しフォーム
    /// </summary>
    private class HiddenForm : Form
    {
        private bool _isInitialized = false;
        public event EventHandler<RawInputEventArgs>? RawInputReceived;
        public event EventHandler<DeviceChangeEventArgs>? DeviceChangeReceived;
        private IntPtr _hDevNotify = IntPtr.Zero;

        public HiddenForm()
        {
            WindowState = FormWindowState.Minimized;
            ShowInTaskbar = false;
            Width = 1;
            Height = 1;
            FormBorderStyle = FormBorderStyle.None;
            StartPosition = FormStartPosition.Manual;
            Location = new Point(-2000, -2000);
        }

        public void Initialize()
        {
            CreateHandle();
            _isInitialized = true;
            RegisterForDeviceNotifications();
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == RawInputConstants.WM_INPUT)
            {
                RawInputReceived?.Invoke(this, new RawInputEventArgs(m.LParam));
            }
            else if (m.Msg == RawInputConstants.WM_INPUT_DEVICE_CHANGE)
            {
                // wParam: GIDC_ARRIVAL(1) / GIDC_REMOVAL(2)
                int change = m.WParam.ToInt32();
                DeviceChangeReceived?.Invoke(this, new DeviceChangeEventArgs(change, m.LParam));
            }
            else if (m.Msg == DeviceNotificationConstants.WM_DEVICECHANGE)
            {
                int wParam = m.WParam.ToInt32();
                if (wParam == DeviceNotificationConstants.DBT_DEVICEARRIVAL ||
                    wParam == DeviceNotificationConstants.DBT_DEVICEREMOVECOMPLETE)
                {
                    try
                    {
                        var hdr = Marshal.PtrToStructure<DEV_BROADCAST_HDR>(m.LParam);
                        if (hdr.dbch_devicetype == DeviceNotificationConstants.DBT_DEVTYP_DEVICEINTERFACE)
                        {
                            // Parse interface detail
                            int baseSize = Marshal.SizeOf<DEV_BROADCAST_DEVICEINTERFACE>();
                            IntPtr pName = IntPtr.Add(m.LParam, baseSize);
                            string? path = Marshal.PtrToStringUni(pName);
                            if (!string.IsNullOrWhiteSpace(path))
                            {
                                bool arrived = (wParam == DeviceNotificationConstants.DBT_DEVICEARRIVAL);
                                // Reuse DeviceChangeEventArgs to propagate with interface path
                                DeviceChangeReceived?.Invoke(this, new DeviceChangeEventArgs(
                                    arrived ? RawInputConstants.GIDC_ARRIVAL : RawInputConstants.GIDC_REMOVAL,
                                    IntPtr.Zero,
                                    path));
                            }
                        }
                    }
                    catch { /* ignore */ }
                }
            }
            base.WndProc(ref m);
        }

        protected override void SetVisibleCore(bool value)
        {
            // 初期化段階では基底クラスの処理を通すが、その後は常に非表示
            if (!_isInitialized && !IsHandleCreated)
            {
                base.SetVisibleCore(value);
            }
            else
            {
                base.SetVisibleCore(false);
            }
        }

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ExStyle |= 0x80; // WS_EX_TOOLWINDOW - タスクバーに表示しない
                cp.ExStyle |= 0x08; // WS_EX_TOPMOST
                return cp;
            }
        }

        private void RegisterForDeviceNotifications()
        {
            try
            {
                var filter = new DEV_BROADCAST_DEVICEINTERFACE
                {
                    dbcc_size = Marshal.SizeOf<DEV_BROADCAST_DEVICEINTERFACE>(),
                    dbcc_devicetype = DeviceNotificationConstants.DBT_DEVTYP_DEVICEINTERFACE,
                    dbcc_reserved = 0,
                    dbcc_classguid = DeviceInterfaceGuids.Mouse
                };

                IntPtr pFilter = Marshal.AllocHGlobal(filter.dbcc_size);
                try
                {
                    Marshal.StructureToPtr(filter, pFilter, false);
                    _hDevNotify = RegisterDeviceNotification(this.Handle, pFilter, DeviceNotificationConstants.DEVICE_NOTIFY_WINDOW_HANDLE);
                }
                finally
                {
                    Marshal.FreeHGlobal(pFilter);
                }
            }
            catch { /* ignore */ }
        }

        protected override void Dispose(bool disposing)
        {
            try
            {
                if (_hDevNotify != IntPtr.Zero)
                {
                    UnregisterDeviceNotification(_hDevNotify);
                    _hDevNotify = IntPtr.Zero;
                }
            }
            catch { /* ignore */ }
            base.Dispose(disposing);
        }
    }

    private class RawInputEventArgs : EventArgs
    {
        public IntPtr LParam { get; }
        public RawInputEventArgs(IntPtr lParam) => LParam = lParam;
    }

    private class DeviceChangeEventArgs : EventArgs
    {
        public int Change { get; }
        public IntPtr LParam { get; }
        public string? InterfacePath { get; }
        public DeviceChangeEventArgs(int change, IntPtr lParam, string? interfacePath = null)
        {
            Change = change;
            LParam = lParam;
            InterfacePath = interfacePath;
        }
    }

    private void OnRawInputDeviceChangeReceived(object? sender, DeviceChangeEventArgs e)
    {
        try
        {
            string deviceId = e.InterfacePath ?? GetDeviceId(e.LParam);
            bool arrived = e.Change == RawInputConstants.GIDC_ARRIVAL;
            System.Diagnostics.Debug.WriteLine($"[RawInput] Device {(arrived ? "ARRIVAL" : "REMOVAL")} : {deviceId}");
            DeviceConnectionChanged?.Invoke(this, new DeviceConnectionChangedEventArgs(deviceId, arrived));
        }
        catch
        {
            // ignore errors for robustness
        }
    }
}
