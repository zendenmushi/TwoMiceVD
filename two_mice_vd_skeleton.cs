using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using Newtonsoft.Json;
using WindowsInput;
using WindowsInput.Native;

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
            try
            {
                // 設定ファイルの読み込み
                _config = ConfigStore.Load();

                // Raw Input Manager の初期化
                try
                {
                    _rawInput = new RawInputManager();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Raw Input の初期化に失敗しました: {ex.Message}\n\n" +
                                    "管理者権限で実行してみてください。", "エラー",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                    throw;
                }

                // 仮想デスクトップコントローラーの初期化
                _vdController = new VirtualDesktopController(_config);
                
                // ポリシーエンジンの初期化
                _policy = new SwitchPolicy(_vdController, _config);

                // タスクトレイUIの初期化
                try
                {
                    _tray = new TrayUI(_config, _rawInput, _policy);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"タスクトレイアイコンの初期化に失敗しました: {ex.Message}", "エラー",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                    throw;
                }

                // Raw input からの移動イベントを処理
                _rawInput.DeviceMoved += OnDeviceMoved;

                // 仮想デスクトップ2枚の存在を前提とするため、チェックは行わない

                // ペアリングメニュー
                _tray.PairingRequested += OnPairingRequested;

                // 初期化完了通知
                _tray?.ShowNotification("TwoMiceVD が開始されました", ToolTipIcon.Info);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"アプリケーションの初期化に失敗しました: {ex.Message}", "致命的エラー",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
                Environment.Exit(1);
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _rawInput?.Dispose();
                _tray?.Dispose();
            }
            base.Dispose(disposing);
        }

        private void OnDeviceMoved(object sender, DeviceMovedEventArgs e)
        {
            _policy.HandleMovement(e.DeviceId, e.DeltaX, e.DeltaY);
        }

        private void OnPairingRequested(object sender, EventArgs e)
        {
            // ペアリング完了後、仮想デスクトップ枚数を再確保
            _vdController.EnsureDesktopCount(_config.VirtualDesktopTargetCount);
            
            // 設定が更新された可能性があるため、ポリシーに反映
            // 必要に応じて追加の初期化処理を行う
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
    public class RawInputManager : IDisposable
    {
        private const int WM_INPUT = 0x00FF;
        private readonly HiddenForm _hiddenForm;

        public event EventHandler<DeviceMovedEventArgs> DeviceMoved;

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
            rid[0].hwndTarget = _hiddenForm.Handle;
            if (!RegisterRawInputDevices(rid, (uint)rid.Length, (uint)Marshal.SizeOf<RAWINPUTDEVICE>()))
            {
                throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());
            }
        }

        private void OnRawInputReceived(object sender, RawInputEventArgs e)
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
            public event EventHandler<RawInputEventArgs> RawInputReceived;

            public HiddenForm()
            {
                WindowState = FormWindowState.Minimized;
                ShowInTaskbar = false;
                Width = 1;
                Height = 1;
                FormBorderStyle = FormBorderStyle.None;
                StartPosition = FormStartPosition.Manual;
                Location = new System.Drawing.Point(-2000, -2000);
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

        #region PInvoke for Raw Input
        [StructLayout(LayoutKind.Sequential)]
        private struct RAWINPUTDEVICE
        {
            public ushort usUsagePage;
            public ushort usUsage;
            public RawInputDeviceFlags dwFlags;
            public IntPtr hwndTarget;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct RAWINPUTHEADER
        {
            public RawInputType dwType;
            public uint dwSize;
            public IntPtr hDevice;
            public IntPtr wParam;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct RAWINPUT
        {
            public RAWINPUTHEADER header;
            public RAWMOUSE data;
        }

        [StructLayout(LayoutKind.Explicit)]
        private struct RAWMOUSE
        {
            [FieldOffset(0)]
            public MOUSE mouse;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct MOUSE
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

        [Flags]
        private enum RawInputDeviceFlags : uint
        {
            INPUTSINK = 0x00000100,
        }

        private enum RawInputType : uint
        {
            RIM_TYPEMOUSE = 0,
            RIM_TYPEKEYBOARD = 1,
            RIM_TYPEHID = 2
        }

        private enum RawInputCommand : uint
        {
            RID_INPUT = 0x10000003,
            RID_HEADER = 0x10000005
        }


        [Flags]
        private enum RawMouseFlags : ushort
        {
            MOUSE_MOVE_RELATIVE = 0x00,
            MOUSE_MOVE_ABSOLUTE = 0x01,
            MOUSE_VIRTUAL_DESKTOP = 0x02,
            MOUSE_ATTRIBUTES_CHANGED = 0x04
        }

        [DllImport("User32.dll", SetLastError = true)]
        private static extern bool RegisterRawInputDevices([In] RAWINPUTDEVICE[] pRawInputDevices, uint uiNumDevices, uint cbSize);

        [DllImport("User32.dll", SetLastError = true)]
        private static extern uint GetRawInputData(IntPtr hRawInput, RawInputCommand uiCommand, IntPtr pData, ref uint pcbSize, uint cbSizeHeader);

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
    /// 仮想デスクトップの切替を管理するクラス。H.InputSimulatorを使用して2枚の仮想デスクトップ間を切り替える。
    /// </summary>
    public class VirtualDesktopController
    {
        private int _currentDesktopIndex = 0;
        private readonly InputSimulator _inputSimulator;
        private readonly Dictionary<int, Point> _desktopCursorPositions = new Dictionary<int, Point>();

        public VirtualDesktopController(ConfigStore config)
        {
            // InputSimulatorのインスタンスを初期化
            _inputSimulator = new InputSimulator();
        }

        public void EnsureDesktopCount(int target)
        {
            // 仮想デスクトップ2枚の存在を前提とするため、何も行わない
            // ユーザーが事前にWin+Ctrl+Dで仮想デスクトップを2枚作成済みと想定
        }

        public void SwitchTo(string desktopId)
        {
            // "VD0" または "VD1" からインデックス 0 または 1 を取得
            if (int.TryParse(desktopId.Replace("VD", ""), out int targetIndex))
            {
                if (targetIndex == _currentDesktopIndex) return;

                // 現在のマウス位置を保存
                SaveCurrentCursorPosition();

                if (targetIndex == 0)
                {
                    SendDesktopLeftShortcut(); // デスクトップ1へ切り替え
                }
                else if (targetIndex == 1)
                {
                    SendDesktopRightShortcut(); // デスクトップ2へ切り替え  
                }

                _currentDesktopIndex = targetIndex;
                
                // 少し待機してからマウス位置を復元（デスクトップ切り替えの完了を待つ）
                System.Threading.Thread.Sleep(100);
                RestoreCursorPosition(targetIndex);
            }
        }

        private void SendDesktopLeftShortcut()
        {
            // Win+Ctrl+左矢印でデスクトップ1へ切り替え
            _inputSimulator.Keyboard.ModifiedKeyStroke(
                new[] { VirtualKeyCode.LWIN, VirtualKeyCode.CONTROL },
                VirtualKeyCode.LEFT
            );
        }

        private void SendDesktopRightShortcut()
        {
            // Win+Ctrl+右矢印でデスクトップ2へ切り替え
            _inputSimulator.Keyboard.ModifiedKeyStroke(
                new[] { VirtualKeyCode.LWIN, VirtualKeyCode.CONTROL },
                VirtualKeyCode.RIGHT
            );
        }

        /// <summary>
        /// 現在のマウス位置を保存する
        /// </summary>
        private void SaveCurrentCursorPosition()
        {
            if (GetCursorPos(out POINT point))
            {
                _desktopCursorPositions[_currentDesktopIndex] = new Point(point.X, point.Y);
            }
        }

        /// <summary>
        /// 指定されたデスクトップのマウス位置を復元する
        /// </summary>
        /// <summary>
        /// 指定されたデスクトップのマウス位置を復元する
        /// </summary>
        private void RestoreCursorPosition(int desktopIndex)
        {
            if (_desktopCursorPositions.TryGetValue(desktopIndex, out Point position))
            {
                // Win32 APIのSetCursorPosを使用してマウス位置を設定
                SetCursorPos(position.X, position.Y);
            }
        }

        #region Win32 API for Cursor Position
        [StructLayout(LayoutKind.Sequential)]
        private struct POINT
        {
            public int X;
            public int Y;
        }

        [DllImport("user32.dll")]
        private static extern bool GetCursorPos(out POINT lpPoint);

        [DllImport("user32.dll")]
        private static extern bool SetCursorPos(int X, int Y);
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
        /// <summary>
        /// ペアリング中の切り替え無効化フラグ
        /// </summary>
        public bool IsPairing { get; set; } = false;

        public SwitchPolicy(VirtualDesktopController controller, ConfigStore config)
        {
            _controller = controller;
            _config = config;
        }

        public void HandleMovement(string deviceId, int dx, int dy)
        {
            // ペアリング中は切り替えを無効化
            if (IsPairing) return;

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
        private readonly RawInputManager _rawInputManager;
        private ContextMenuStrip _menu;
        private PairingDialog _pairingDialog;
        private SettingsDialog _settingsDialog;
        private readonly SwitchPolicy _switchPolicy;
        
        public event EventHandler PairingRequested;

        public TrayUI(ConfigStore config, RawInputManager rawInputManager, SwitchPolicy switchPolicy)
        {
            _config = config;
            _rawInputManager = rawInputManager;
            _switchPolicy = switchPolicy;
            _menu = BuildContextMenu(); // 強参照を保持してGCによる解放を防ぐ
            _icon = new NotifyIcon
            {
                Text = "TwoMiceVD - 2マウス仮想デスクトップ切替",
                Icon = System.Drawing.SystemIcons.Application,
                Visible = true,
                ContextMenuStrip = _menu
            };
            
            _icon.DoubleClick += (s, e) => ShowSettings();

            // Windows 10/11 で右クリックメニューが稀に出ない対策として手動表示も実装
            _icon.MouseUp += (s, e) =>
            {
                try
                {
                    if (e.Button == MouseButtons.Right)
                    {
                        // タスクトレイ位置にメニューを表示
                        _menu?.Show(System.Windows.Forms.Control.MousePosition);
                    }
                    else if (e.Button == MouseButtons.Left)
                    {
                        // 左クリックでもメニューを出したい場合は以下を有効化
                        // _menu?.Show(System.Windows.Forms.Control.MousePosition);
                    }
                }
                catch { /* 例外は握り潰して安定性優先 */ }
            };
        }

        private ContextMenuStrip BuildContextMenu()
        {
            var menu = new ContextMenuStrip();

            var pairItem = new ToolStripMenuItem("ペアリング開始")
            {
                ToolTipText = "マウスA/Bの識別を再設定します"
            };
            pairItem.Click += (s, e) => StartPairing();
            menu.Items.Add(pairItem);

            menu.Items.Add(new ToolStripSeparator());

            var swapItem = new ToolStripMenuItem("マウス割当反転")
            {
                Checked = _config.SwapBindings,
                CheckOnClick = true,
                ToolTipText = "マウスA/Bの割り当てを入れ替えます"
            };
            swapItem.Click += (s, e) => 
            {
                _config.SwapBindings = swapItem.Checked;
                _config.Save();
            };
            menu.Items.Add(swapItem);

            var settingsItem = new ToolStripMenuItem("設定")
            {
                ToolTipText = "感度、クールダウンなどを設定します"
            };
            settingsItem.Click += (s, e) => ShowSettings();
            menu.Items.Add(settingsItem);

            var startupItem = new ToolStripMenuItem("起動時に自動開始")
            {
                Checked = _config.StartupRun,
                CheckOnClick = true,
                ToolTipText = "Windows起動時にアプリを自動起動します"
            };
            startupItem.Click += (s, e) =>
            {
                _config.StartupRun = startupItem.Checked;
                SetStartupRegistration(startupItem.Checked);
                _config.Save();
            };
            menu.Items.Add(startupItem);

            menu.Items.Add(new ToolStripSeparator());

            var aboutItem = new ToolStripMenuItem("TwoMiceVD について");
            aboutItem.Click += (s, e) => ShowAbout();
            menu.Items.Add(aboutItem);

            var exitItem = new ToolStripMenuItem("終了");
            exitItem.Click += (s, e) => Application.Exit();
            menu.Items.Add(exitItem);

            return menu;
        }

        private void StartPairing()
        {
            if (_pairingDialog == null || _pairingDialog.IsDisposed)
            {
                _pairingDialog = new PairingDialog(_config, _rawInputManager);
            }
            
            // ペアリング開始時に切り替えを無効化
            _switchPolicy.IsPairing = true;
            
            _pairingDialog.ShowDialog();
            
            // ペアリング終了時に切り替えを有効化
            _switchPolicy.IsPairing = false;
            
            PairingRequested?.Invoke(this, EventArgs.Empty);
        }

        private void ShowSettings()
        {
            if (_settingsDialog == null || _settingsDialog.IsDisposed)
            {
                _settingsDialog = new SettingsDialog(_config);
            }
            
            _settingsDialog.ShowDialog();
        }

        private void ShowAbout()
        {
            MessageBox.Show(
                "TwoMiceVD v1.0\n\n" +
                "2台の物理マウスを個別に認識し、各マウスに紐づく\n" +
                "仮想デスクトップへ自動で切り替えるアプリケーション\n\n" +
                "Windows 11 22H2 以降対応",
                "TwoMiceVD について",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            );
        }

        private void SetStartupRegistration(bool enable)
        {
            const string keyName = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run";
            const string valueName = "TwoMiceVD";

            try
            {
                using (var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(keyName, true))
                {
                    if (enable)
                    {
                        string exePath = System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName;
                        key?.SetValue(valueName, exePath);
                    }
                    else
                    {
                        key?.DeleteValue(valueName, false);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"起動設定の変更に失敗しました: {ex.Message}", "エラー", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        public void ShowNotification(string message, ToolTipIcon icon = ToolTipIcon.Info)
        {
            _icon.ShowBalloonTip(3000, "TwoMiceVD", message, icon);
        }

        public void Dispose()
        {
            if (_icon != null)
            {
                _icon.Visible = false;
                _icon.Dispose();
            }
            _pairingDialog?.Dispose();
            _settingsDialog?.Dispose();
            _menu?.Dispose();
        }
    }

    /// <summary>
    /// マウスペアリング用ダイアログ
    /// </summary>
    public partial class PairingDialog : Form
    {
        private readonly ConfigStore _config;
        private readonly RawInputManager _sharedRawInput;
        private readonly Dictionary<string, string> _deviceNames = new Dictionary<string, string>();
        private bool _isCapturingA = false;
        private bool _isCapturingB = false;
        private Label _instructionLabel;
        private Button _captureAButton;
        private Button _captureBButton;
        private Button _completeButton;
        private Label _statusLabel;

        public PairingDialog(ConfigStore config, RawInputManager sharedRawInput)
        {
            _config = config;
            _sharedRawInput = sharedRawInput;
            _sharedRawInput.DeviceMoved += OnDeviceMoved;
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            Text = "マウスペアリング";
            Size = new System.Drawing.Size(400, 300);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;

            _instructionLabel = new Label
            {
                Text = "各マウスを順番に動かして識別してください",
                Location = new System.Drawing.Point(20, 20),
                Size = new System.Drawing.Size(360, 40),
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            };
            Controls.Add(_instructionLabel);

            _captureAButton = new Button
            {
                Text = "マウスA をキャプチャ",
                Location = new System.Drawing.Point(50, 80),
                Size = new System.Drawing.Size(120, 30)
            };
            _captureAButton.Click += (s, e) => StartCapture("A");
            Controls.Add(_captureAButton);

            _captureBButton = new Button
            {
                Text = "マウスB をキャプチャ",
                Location = new System.Drawing.Point(230, 80),
                Size = new System.Drawing.Size(120, 30)
            };
            _captureBButton.Click += (s, e) => StartCapture("B");
            Controls.Add(_captureBButton);

            _statusLabel = new Label
            {
                Text = "",
                Location = new System.Drawing.Point(20, 130),
                Size = new System.Drawing.Size(360, 60),
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            };
            Controls.Add(_statusLabel);

            _completeButton = new Button
            {
                Text = "完了",
                Location = new System.Drawing.Point(160, 220),
                Size = new System.Drawing.Size(80, 30),
                Enabled = false
            };
            _completeButton.Click += (s, e) => Complete();
            Controls.Add(_completeButton);
        }

        private void StartCapture(string deviceKey)
        {
            _isCapturingA = (deviceKey == "A");
            _isCapturingB = (deviceKey == "B");
            _statusLabel.Text = $"{deviceKey} を識別するため、そのマウスを動かしてください...";
            
            _captureAButton.Enabled = !_isCapturingA;
            _captureBButton.Enabled = !_isCapturingB;
        }

        private void OnDeviceMoved(object sender, DeviceMovedEventArgs e)
        {
            if (_isCapturingA)
            {
                _deviceNames["A"] = e.DeviceId;
                _isCapturingA = false;
                _statusLabel.Text = "マウスA の識別が完了しました";
                _captureAButton.Text = "マウスA ✓";
                _captureAButton.Enabled = true;
                _captureBButton.Enabled = true;
            }
            else if (_isCapturingB)
            {
                _deviceNames["B"] = e.DeviceId;
                _isCapturingB = false;
                _statusLabel.Text = "マウスB の識別が完了しました";
                _captureBButton.Text = "マウスB ✓";
                _captureAButton.Enabled = true;
                _captureBButton.Enabled = true;
            }

            if (_deviceNames.ContainsKey("A") && _deviceNames.ContainsKey("B"))
            {
                _completeButton.Enabled = true;
                _statusLabel.Text = "両方のマウスが識別されました。完了ボタンを押してください。";
            }
        }

        private void Complete()
        {
            foreach (var device in _deviceNames)
            {
                _config.SetDeviceBinding(device.Key, device.Value, 0, 0);
                _config.SetDesktopBinding(device.Key, $"VD{(device.Key == "A" ? 0 : 1)}");
            }
            _config.Save();
            
            MessageBox.Show("マウスのペアリングが完了しました", "完了", 
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            Close();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _sharedRawInput.DeviceMoved -= OnDeviceMoved;
            }
            base.Dispose(disposing);
        }
    }

    /// <summary>
    /// 設定ダイアログ
    /// </summary>
    public partial class SettingsDialog : Form
    {
        private readonly ConfigStore _config;
        private TrackBar _thresholdTrackBar;
        private TrackBar _hysteresisTrackBar;
        private Label _thresholdLabel;
        private Label _hysteresisLabel;
        private Button _okButton;
        private Button _cancelButton;

        public SettingsDialog(ConfigStore config)
        {
            _config = config;
            InitializeComponent();
            LoadSettings();
        }

        private void InitializeComponent()
        {
            Text = "設定";
            Size = new System.Drawing.Size(400, 250);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;

            var thresholdTitleLabel = new Label
            {
                Text = "感度（移動しきい値）",
                Location = new System.Drawing.Point(20, 20),
                Size = new System.Drawing.Size(150, 20)
            };
            Controls.Add(thresholdTitleLabel);

            _thresholdLabel = new Label
            {
                Location = new System.Drawing.Point(300, 20),
                Size = new System.Drawing.Size(50, 20),
                TextAlign = System.Drawing.ContentAlignment.MiddleRight
            };
            Controls.Add(_thresholdLabel);

            _thresholdTrackBar = new TrackBar
            {
                Location = new System.Drawing.Point(20, 45),
                Size = new System.Drawing.Size(330, 45),
                Minimum = 1,
                Maximum = 20,
                TickFrequency = 1
            };
            _thresholdTrackBar.ValueChanged += (s, e) => 
                _thresholdLabel.Text = _thresholdTrackBar.Value.ToString();
            Controls.Add(_thresholdTrackBar);

            var hysteresisTitleLabel = new Label
            {
                Text = "クールダウン時間（ミリ秒）",
                Location = new System.Drawing.Point(20, 100),
                Size = new System.Drawing.Size(150, 20)
            };
            Controls.Add(hysteresisTitleLabel);

            _hysteresisLabel = new Label
            {
                Location = new System.Drawing.Point(300, 100),
                Size = new System.Drawing.Size(50, 20),
                TextAlign = System.Drawing.ContentAlignment.MiddleRight
            };
            Controls.Add(_hysteresisLabel);

            _hysteresisTrackBar = new TrackBar
            {
                Location = new System.Drawing.Point(20, 125),
                Size = new System.Drawing.Size(330, 45),
                Minimum = 100,
                Maximum = 1000,
                TickFrequency = 100
            };
            _hysteresisTrackBar.ValueChanged += (s, e) => 
                _hysteresisLabel.Text = _hysteresisTrackBar.Value.ToString();
            Controls.Add(_hysteresisTrackBar);

            _okButton = new Button
            {
                Text = "OK",
                Location = new System.Drawing.Point(220, 180),
                Size = new System.Drawing.Size(75, 25),
                DialogResult = DialogResult.OK
            };
            _okButton.Click += (s, e) => SaveSettings();
            Controls.Add(_okButton);

            _cancelButton = new Button
            {
                Text = "キャンセル",
                Location = new System.Drawing.Point(305, 180),
                Size = new System.Drawing.Size(75, 25),
                DialogResult = DialogResult.Cancel
            };
            Controls.Add(_cancelButton);

            AcceptButton = _okButton;
            CancelButton = _cancelButton;
        }

        private void LoadSettings()
        {
            _thresholdTrackBar.Value = _config.ThresholdMovement;
            _hysteresisTrackBar.Value = _config.HysteresisMs;
            _thresholdLabel.Text = _config.ThresholdMovement.ToString();
            _hysteresisLabel.Text = _config.HysteresisMs.ToString();
        }

        private void SaveSettings()
        {
            _config.ThresholdMovement = _thresholdTrackBar.Value;
            _config.HysteresisMs = _hysteresisTrackBar.Value;
            _config.Save();
        }
    }

    /// <summary>
    /// 設定ファイルを読み書きするクラス。各種しきい値やマッピングを保持する。
    /// </summary>
    public class ConfigStore
    {
        private static readonly string ConfigDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "TwoMiceVD");
        private static readonly string ConfigFilePath = Path.Combine(ConfigDirectory, "config.json");

        public int ThresholdMovement { get; set; } = 5;
        public int HysteresisMs { get; set; } = 400;
        public int MaxSwitchPerSec { get; set; } = 2;
        public int VirtualDesktopTargetCount { get; set; } = 2;
        public bool PreferInternalApi { get; set; } = true;
        public bool StartupRun { get; set; } = false;
        public bool SwapBindings { get; set; } = false;

        public Dictionary<string, DeviceInfo> Devices { get; set; } = new Dictionary<string, DeviceInfo>();
        public Dictionary<string, string> Bindings { get; set; } = new Dictionary<string, string>();
        public VirtualDesktopInfo VirtualDesktops { get; set; } = new VirtualDesktopInfo();

        public static ConfigStore Load()
        {
            try
            {
                if (File.Exists(ConfigFilePath))
                {
                    string json = File.ReadAllText(ConfigFilePath);
                    var config = JsonConvert.DeserializeObject<ConfigData>(json);
                    return FromConfigData(config);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"設定ファイルの読み込みに失敗しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            return new ConfigStore();
        }

        public void Save()
        {
            try
            {
                Directory.CreateDirectory(ConfigDirectory);
                var configData = ToConfigData();
                string json = JsonConvert.SerializeObject(configData, Formatting.Indented);
                File.WriteAllText(ConfigFilePath, json);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"設定ファイルの保存に失敗しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public string GetDesktopIdForDevice(string deviceId)
        {
            if (deviceId == null) return null;

            // デバイスID から デバイス識別子（A/B）を取得
            string deviceKey = null;
            foreach (var device in Devices)
            {
                if (device.Value.DevicePath == deviceId)
                {
                    deviceKey = device.Key;
                    break;
                }
            }

            if (deviceKey == null) return null;

            // 割当の反転が有効な場合は反転させる
            if (SwapBindings)
            {
                deviceKey = deviceKey == "A" ? "B" : "A";
            }

            // デバイス識別子から仮想デスクトップIDを取得
            if (Bindings.TryGetValue(deviceKey, out string desktopKey))
            {
                if (VirtualDesktops.Mode == "GUID" && VirtualDesktops.Ids.TryGetValue(desktopKey, out string guid))
                {
                    return guid;
                }
                return desktopKey; // インデックス形式の場合
            }

            return null;
        }

        public void SetDeviceBinding(string deviceKey, string devicePath, int vid, int pid)
        {
            Devices[deviceKey] = new DeviceInfo { DevicePath = devicePath, Vid = vid, Pid = pid };
        }

        public void SetDesktopBinding(string deviceKey, string desktopKey, string desktopGuid = null)
        {
            Bindings[deviceKey] = desktopKey;
            if (!string.IsNullOrEmpty(desktopGuid))
            {
                VirtualDesktops.Mode = "GUID";
                VirtualDesktops.Ids[desktopKey] = desktopGuid;
            }
        }

        private ConfigData ToConfigData()
        {
            return new ConfigData
            {
                devices = Devices,
                bindings = Bindings,
                virtual_desktops = VirtualDesktops,
                @switch = new SwitchConfig
                {
                    threshold = ThresholdMovement,
                    hysteresis_ms = HysteresisMs,
                    max_per_sec = MaxSwitchPerSec
                },
                ui = new UiConfig
                {
                    startup_run = StartupRun,
                    swap_bindings = SwapBindings
                },
                impl = new ImplementationConfig
                {
                    prefer_internal_api = PreferInternalApi
                }
            };
        }

        private static ConfigStore FromConfigData(ConfigData data)
        {
            var config = new ConfigStore();
            
            if (data.devices != null)
                config.Devices = data.devices;
            
            if (data.bindings != null)
                config.Bindings = data.bindings;
            
            if (data.virtual_desktops != null)
                config.VirtualDesktops = data.virtual_desktops;
            
            if (data.@switch != null)
            {
                config.ThresholdMovement = data.@switch.threshold;
                config.HysteresisMs = data.@switch.hysteresis_ms;
                config.MaxSwitchPerSec = data.@switch.max_per_sec;
            }
            
            if (data.ui != null)
            {
                config.StartupRun = data.ui.startup_run;
                config.SwapBindings = data.ui.swap_bindings;
            }
            
            if (data.impl != null)
            {
                config.PreferInternalApi = data.impl.prefer_internal_api;
            }
            
            return config;
        }

        public class DeviceInfo
        {
            [JsonProperty("device_path")]
            public string DevicePath { get; set; }
            
            [JsonProperty("vid")]
            public int Vid { get; set; }
            
            [JsonProperty("pid")]
            public int Pid { get; set; }
        }

        public class VirtualDesktopInfo
        {
            [JsonProperty("mode")]
            public string Mode { get; set; } = "GUID";
            
            [JsonProperty("ids")]
            public Dictionary<string, string> Ids { get; set; } = new Dictionary<string, string>();
            
            [JsonProperty("target_count")]
            public int TargetCount { get; set; } = 2;
        }

        private class ConfigData
        {
            public Dictionary<string, DeviceInfo> devices { get; set; }
            public Dictionary<string, string> bindings { get; set; }
            public VirtualDesktopInfo virtual_desktops { get; set; }
            public SwitchConfig @switch { get; set; }
            public UiConfig ui { get; set; }
            public ImplementationConfig impl { get; set; }
        }

        private class SwitchConfig
        {
            public int threshold { get; set; }
            public int hysteresis_ms { get; set; }
            public int max_per_sec { get; set; }
        }

        private class UiConfig
        {
            public bool startup_run { get; set; }
            public bool swap_bindings { get; set; }
        }

        private class ImplementationConfig
        {
            public bool prefer_internal_api { get; set; }
        }
    }
}
