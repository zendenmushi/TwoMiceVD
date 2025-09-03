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
    private const int WM_INPUT = 0x00FF;
    private readonly HiddenForm? _hiddenForm;

    public event EventHandler<DeviceMovedEventArgs>? DeviceMoved;

    public RawInputManager()
    {
        _hiddenForm = new HiddenForm();
        _hiddenForm.RawInputReceived += OnRawInputReceived;
        
        // フォームのハンドル作成を明示的に実行
        _hiddenForm.Initialize();
        
        // ハンドルが作成されてからRaw Input を登録
        if (!_hiddenForm.IsHandleCreated)
        {
            throw new InvalidOperationException("HiddenFormのハンドル作成に失敗しました");
        }
        
        RegisterRawInput();
    }

    private void RegisterRawInput()
    {
        RAWINPUTDEVICE[] rid = new RAWINPUTDEVICE[1];
        rid[0].usUsagePage = 0x01;
        rid[0].usUsage = 0x02; // マウス
        rid[0].dwFlags = RawInputDeviceFlags.INPUTSINK;
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
                    int deltaX = rawInput.data.mouse.lLastX;
                    int deltaY = rawInput.data.mouse.lLastY;

                    if (deltaX != 0 || deltaY != 0)
                    {
                        DeviceMoved?.Invoke(this, new DeviceMovedEventArgs(deviceId, deltaX, deltaY));
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
        // ハンドル値をそのまま使用して文字化け問題を回避
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
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_INPUT)
            {
                RawInputReceived?.Invoke(this, new RawInputEventArgs(m.LParam));
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
    }

    private class RawInputEventArgs : EventArgs
    {
        public IntPtr LParam { get; }
        public RawInputEventArgs(IntPtr lParam) => LParam = lParam;
    }
}