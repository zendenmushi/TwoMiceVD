using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace TwoMiceVD
{
    /// <summary>
    /// エントリポイントとなるコンテキストクラス。RawInputManager, VirtualDesktopController, SwitchPolicy, TrayUI, ConfigStore
    /// を初期化し、イベントの配線を行います。
    /// </summary>
    class Program : ApplicationContext
    {
        private readonly RawInputManager _rawInput;
        private readonly VirtualDesktopController _vdController;
        private readonly SwitchPolicy _policy;
        private readonly TrayUI _tray;
        private readonly ConfigStore _config;

        public Program()
        {
            _config = ConfigStore.Load();
            _rawInput = new RawInputManager();
            _vdController = new VirtualDesktopController(_config);
            _policy = new SwitchPolicy(_vdController, _config);
            _tray = new TrayUI(_config);

            // Raw input からの移動イベントを処理
            _rawInput.DeviceMoved += OnDeviceMoved;

            // 仮想デスクトップ枚数を確保
            _vdController.EnsureDesktopCount(_config.VirtualDesktopTargetCount);

            // ペアリングメニュー
            _tray.PairingRequested += OnPairingRequested;
        }

        private void OnDeviceMoved(object sender, DeviceMovedEventArgs e)
        {
            _policy.HandleMovement(e.DeviceId, e.DeltaX, e.DeltaY);
        }

        private void OnPairingRequested(object sender, EventArgs e)
        {
            // TODO: ペアリングダイアログを実装してマウスA/Bを識別し、ConfigStore に保存する
        }

        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Program());
        }
    }

    /// <summary>
    /// Raw Input API を使ってマウス入力をバックグラウンドで受信し、移動イベントを発行する。
    /// </summary>
    public class RawInputManager
    {
        public event EventHandler<DeviceMovedEventArgs> DeviceMoved;

        public RawInputManager()
        {
            RegisterRawInput();
        }

        private void RegisterRawInput()
        {
            RAWINPUTDEVICE[] rid = new RAWINPUTDEVICE[1];
            rid[0].usUsagePage = 0x01;
            rid[0].usUsage = 0x02; // マウス
            rid[0].dwFlags = RawInputDeviceFlags.INPUTSINK;
            rid[0].hwndTarget = IntPtr.Zero;
            if (!RegisterRawInputDevices(rid, (uint)rid.Length, (uint)Marshal.SizeOf<RAWINPUTDEVICE>()))
            {
                throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());
            }
        }

        // TODO: 隠しフォームで WndProc をオーバーライドし、WM_INPUT を処理して DeviceMoved イベントを発火させる

        #region PInvoke for Raw Input
        [StructLayout(LayoutKind.Sequential)]
        private struct RAWINPUTDEVICE
        {
            public ushort usUsagePage;
            public ushort usUsage;
            public RawInputDeviceFlags dwFlags;
            public IntPtr hwndTarget;
        }

        [Flags]
        private enum RawInputDeviceFlags : uint
        {
            INPUTSINK = 0x00000100,
        }

        [DllImport("User32.dll", SetLastError = true)]
        private static extern bool RegisterRawInputDevices([In] RAWINPUTDEVICE[] pRawInputDevices, uint uiNumDevices, uint cbSize);
        #endregion
    }

    /// <summary>
    /// DeviceId と移動量を通知するイベント引数。
    /// </summary>
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

    /// <summary>
    /// 仮想デスクトップの列挙・生成・切替を管理するクラス。主に COM API を利用し、失敗時に SendInput 方式へフォールバックする。
    /// </summary>
    public class VirtualDesktopController
    {
        private readonly ConfigStore _config;

        public VirtualDesktopController(ConfigStore config)
        {
            _config = config;
            // TODO: COM インターフェイスを初期化
        }

        public void EnsureDesktopCount(int target)
        {
            // TODO: 現在のデスクトップ枚数を調査し、不足する場合は追加する
        }

        public void SwitchTo(string desktopId)
        {
            // TODO: 指定された GUID またはインデックスの仮想デスクトップへ切り替える
            // COM API が失敗した場合は SendInput を用いてフォールバックする
        }

        // TODO: 仮想デスクトップの列挙や GUID 取得などを提供するメソッドを実装

        #region 仮想デスクトップ COM インターフェイスの宣言（抜粋）
        // 以下は必要なインターフェイスの空宣言であり、メソッドシグネチャを定義した上で実装時に利用する
        [ComImport, Guid("C179334C-4295-40D3-BEA1-C654D965605A")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IVirtualDesktop
        {
            // TODO: 必要なメソッドを宣言
        }

        // 例: IVirtualDesktopManager は公開 API なので本格実装で利用可能
        [ComImport, Guid("A5CD92FF-29BE-454C-8D04-D82879FB3F1B")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IVirtualDesktopManager
        {
            int IsWindowOnCurrentVirtualDesktop(IntPtr topLevelWindow, out bool onCurrentDesktop);
            int GetWindowDesktopId(IntPtr topLevelWindow, out Guid desktopId);
            int MoveWindowToDesktop(IntPtr topLevelWindow, Guid desktopId);
        }

        #endregion
    }

    /// <summary>
    /// 入力の累積移動量から閾値とヒステリシスに基づいて仮想デスクトップの切替可否を判断する。
    /// </summary>
    public class SwitchPolicy
    {
        private readonly VirtualDesktopController _controller;
        private readonly ConfigStore _config;
        private DateTime _lastSwitchTime = DateTime.MinValue;
        private readonly Dictionary<string, int> _movementBuckets = new Dictionary<string, int>();

        public SwitchPolicy(VirtualDesktopController controller, ConfigStore config)
        {
            _controller = controller;
            _config = config;
        }

        public void HandleMovement(string deviceId, int dx, int dy)
        {
            int move = Math.Abs(dx) + Math.Abs(dy);
            if (move <= 0) return;

            if (!_movementBuckets.ContainsKey(deviceId))
                _movementBuckets[deviceId] = 0;

            _movementBuckets[deviceId] += move;

            TimeSpan sinceLast = DateTime.Now - _lastSwitchTime;
            if (sinceLast.TotalMilliseconds < _config.HysteresisMs)
                return;

            if (_movementBuckets[deviceId] >= _config.ThresholdMovement)
            {
                string target = _config.GetDesktopIdForDevice(deviceId);
                if (!string.IsNullOrEmpty(target))
                {
                    _controller.SwitchTo(target);
                    _lastSwitchTime = DateTime.Now;
                    _movementBuckets.Clear();
                }
            }
        }
    }

    /// <summary>
    /// システムトレイに常駐し、コンテキストメニューとペアリングリクエストイベントを提供する。
    /// </summary>
    public class TrayUI : IDisposable
    {
        private readonly NotifyIcon _icon;
        private readonly ConfigStore _config;
        public event EventHandler PairingRequested;

        public TrayUI(ConfigStore config)
        {
            _config = config;
            _icon = new NotifyIcon
            {
                Text = "TwoMiceVD",
                Icon = System.Drawing.SystemIcons.Application,
                Visible = true,
                ContextMenuStrip = BuildContextMenu()
            };
        }

        private ContextMenuStrip BuildContextMenu()
        {
            var menu = new ContextMenuStrip();
            var pairItem = new ToolStripMenuItem("ペアリング開始");
            pairItem.Click += (s, e) => PairingRequested?.Invoke(this, EventArgs.Empty);
            menu.Items.Add(pairItem);
            // TODO: 割当反転、感度、クールダウン、起動時自動開始、ログ表示などのメニューを追加

            var exitItem = new ToolStripMenuItem("終了");
            exitItem.Click += (s, e) => Application.Exit();
            menu.Items.Add(exitItem);
            return menu;
        }

        public void Dispose()
        {
            _icon.Dispose();
        }
    }

    /// <summary>
    /// 設定ファイルを読み書きするクラス。各種しきい値やマッピングを保持する。
    /// </summary>
    public class ConfigStore
    {
        public int ThresholdMovement { get; set; } = 5;
        public int HysteresisMs { get; set; } = 400;
        public int MaxSwitchPerSec { get; set; } = 2;
        public int VirtualDesktopTargetCount { get; set; } = 2;
        public bool PreferInternalApi { get; set; } = true;
        // TODO: デバイス識別子と仮想デスクトップ GUID/インデックスのマッピングを保持

        public static ConfigStore Load()
        {
            // TODO: 設定ファイルを読み込む。存在しない場合はデフォルト値で初期化
            return new ConfigStore();
        }

        public void Save()
        {
            // TODO: 設定をファイルに書き込む
        }

        public string GetDesktopIdForDevice(string deviceId)
        {
            // TODO: デバイスIDに対応するデスクトップ GUID/インデックスを返す
            return null;
        }
    }
}
