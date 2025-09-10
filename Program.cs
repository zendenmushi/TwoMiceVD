using System;
using System.Threading.Tasks;
using System.Windows.Forms;
using TwoMiceVD.Core;
using TwoMiceVD.Input;
using TwoMiceVD.UI;
using TwoMiceVD.Configuration;
using TwoMiceVD.Interop;

namespace TwoMiceVD;

internal class Program : ApplicationContext
{
    private readonly RawInputManager? _rawInput;
    private readonly VirtualDesktopController? _vdController;
    private readonly SwitchPolicy? _policy;
    private readonly TrayUI? _tray;
    private readonly ConfigStore? _config;
    private readonly MousePositionManager? _mousePositionManager;
    private readonly System.Collections.Generic.Dictionary<string, int> _buttonsDownCountByDevice = new();
    private bool _autoSwitchTemporarilyDisabled = false;
    private readonly System.Collections.Generic.HashSet<string> _connectedDevicePaths = new(System.StringComparer.OrdinalIgnoreCase);
    
    // 直近で動いているマウス（アクティブ）の判定用
    private string? _activeDeviceId;
    private DateTime _activeUntil = DateTime.MinValue;
    private int _activeHoldMs = 150; // 連続動作の猶予時間（設定連動）
    
    // 初期化・ペアリング制御フラグ
    private bool _isInitializing = true; // 初期化中フラグ（コンストラクタで開始）
    private bool _isPairing = false; // ペアリング中フラグ

    public Program()
    {
        try
        {
#if DEBUG
            // デバッグビルドの場合、Debug出力をコンソールにも表示
            System.Diagnostics.Trace.Listeners.Add(new System.Diagnostics.TextWriterTraceListener(Console.Out));
            System.Diagnostics.Trace.AutoFlush = true;
#endif
            
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
            
            // マウス位置管理の初期化
            _mousePositionManager = new MousePositionManager();

            // アクティブ猶予時間を設定から反映
            _activeHoldMs = Math.Max(50, Math.Min(1000, _config.ActiveHoldMs));

            // タスクトレイUIの初期化
            try
            {
                _tray = new TrayUI(_config, _rawInput, _policy, _vdController);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"タスクトレイアイコンの初期化に失敗しました: {ex.Message}", "エラー",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw;
            }

            // Raw input からの移動イベントを処理
            _rawInput.DeviceMoved += OnDeviceMoved;
            
            // デバイス接続状態確認（オプション）
            if (_config.EnableDeviceConnectionCheck)
            {
                // 起動時の接続状態を列挙して初期化
                try
                {
                    var initial = _rawInput.GetCurrentlyConnectedMouseDevicePaths();
                    foreach (var path in initial)
                    {
                        _connectedDevicePaths.Add(path);
                    }
                }
                catch { /* 列挙失敗は無視 */ }

                // 接続変更通知を購読
                _rawInput.DeviceConnectionChanged += OnDeviceConnectionChanged;
            }
            
            // 初期状態での一時停止/再開を反映
            EvaluateAutoSwitchSuspension();

            // 仮想デスクトップ2枚の存在を前提とするため、チェックは行わない

            // ペアリングメニュー
            _tray.PairingRequested += OnPairingRequested;
            _tray.PairingStarted += OnPairingStarted;
            _tray.PairingCompleted += OnPairingCompleted;

            // 初期化完了通知
            _tray?.ShowNotification("TwoMiceVD が開始されました", ToolTipIcon.Info);
            
            // 起動時のGUID検証とマーカー初期化を順序立てて実行
            _ = Task.Run(async () => await ValidateStartupAndInitializeMarkersAsync());
        }
        catch (Exception ex)
        {
            MessageBox.Show($"アプリケーションの初期化に失敗しました: {ex.Message}", "致命的エラー",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
            Environment.Exit(1);
        }
    }

    /// <summary>
    /// 起動時の保存されたGUID検証
    /// </summary>
    private async Task ValidateStartupGuidsAsync()
    {
        try
        {
            System.Diagnostics.Debug.WriteLine("[Program] GUID検証を開始中...");
            
            if (_config == null || _vdController == null || _tray == null)
            {
                System.Diagnostics.Debug.WriteLine("[Program] 必要なコンポーネントが初期化されていません");
                return;
            }

            // GUID モードでない場合はスキップ
            if (_config.VirtualDesktops.Mode != "GUID")
            {
                System.Diagnostics.Debug.WriteLine("[Program] GUIDモードではないため検証をスキップ");
                return;
            }

            var validationResult = await _config.ValidateStoredGuidsAsync(_vdController);
            if (!validationResult.IsValid)
            {
                System.Diagnostics.Debug.WriteLine($"[Program] GUID検証失敗: {validationResult.Message}");
                // 再ペアリングを促すバルーンは表示する
                _tray.ShowNotification(
                    string.IsNullOrWhiteSpace(validationResult.Message)
                        ? "保存された仮想デスクトップが見つかりません。ペアリングをやり直してください。"
                        : validationResult.Message,
                    ToolTipIcon.Warning
                );
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("[Program] GUID検証が成功しました");
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[Program] GUID検証中にエラー: {ex.Message}");
            
            // デバッグ情報のみ出力し、ユーザーには通知しない
            // マーカーシステムが動作すれば問題ないため
        }
    }

    /// <summary>
    /// GUID検証とマーカーシステム初期化を順序立てて実行
    /// </summary>
    private async Task ValidateStartupAndInitializeMarkersAsync()
    {
        try
        {
            System.Diagnostics.Debug.WriteLine("[Program] GUID検証を開始...");
            
            // Step 1: GUID検証を実行
            await ValidateStartupGuidsAsync();
            
            // Step 2: GUID検証完了後、少し待機してからマーカー初期化
            await Task.Delay(1000); // デスクトップ構成が安定するまで待機
            
            System.Diagnostics.Debug.WriteLine("[Program] マーカーシステムの初期化を開始...");
            
            // Step 3: マーカーシステムを初期化
            await InitializeMarkersAsync();
            
            // Step 4: 初期化完了 - マウス入力処理を有効化
            _isInitializing = false;
            System.Diagnostics.Debug.WriteLine("[Program] 初期化完了 - マウス入力処理を有効化");
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[Program] 統合初期化中にエラー: {ex.Message}");
            _tray?.ShowNotification(
                $"システム初期化中にエラーが発生しました: {ex.Message}",
                ToolTipIcon.Warning
            );
            
            // エラーが発生しても初期化フラグをリセット
            _isInitializing = false;
            System.Diagnostics.Debug.WriteLine("[Program] 初期化エラー後もマウス入力処理を有効化");
        }
    }

    /// <summary>
    /// マーカーシステムの初期化
    /// </summary>
    private async Task InitializeMarkersAsync()
    {
        try
        {
            System.Diagnostics.Debug.WriteLine("[Program] マーカーシステムの初期化を開始...");
            
            // 少し待機してからマーカー初期化を開始（デスクトップ構成安定化のため）
            await Task.Delay(500);
            
            if (_vdController == null)
            {
                System.Diagnostics.Debug.WriteLine("[Program] VirtualDesktopControllerが初期化されていません");
                return;
            }

            // 前提条件をチェック
            System.Diagnostics.Debug.WriteLine("[Program] マーカー初期化の前提条件をチェック中...");
            var (canInitialize, reason) = await _vdController.CheckMarkerInitializationPrerequisitesAsync();
            
            if (!canInitialize)
            {
                System.Diagnostics.Debug.WriteLine($"[Program] 前提条件チェック失敗: {reason}");
                // 情報バルーンは出さない（ログのみ）
                return;
            }

            System.Diagnostics.Debug.WriteLine("[Program] 前提条件チェック成功、マーカー初期化を実行...");
            bool success = await _vdController.InitializeMarkersAsync();
            
            if (success)
            {
                System.Diagnostics.Debug.WriteLine("[Program] マーカーシステムの初期化が完了しました");
                // 成功時の情報バルーンは出さない
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("[Program] マーカーシステムの初期化に失敗しました");
                _tray?.ShowNotification(
                    "高速デスクトップ判定システムの初期化に失敗しました\n通常モードで動作します",
                    ToolTipIcon.Warning
                );
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[Program] マーカー初期化中にエラー: {ex.Message}");
            System.Diagnostics.Debug.WriteLine($"[Program] エラー詳細: {ex}");
            _tray?.ShowNotification(
                $"マーカー初期化中にエラーが発生しました: {ex.Message}",
                ToolTipIcon.Warning
            );
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

    private void OnDeviceMoved(object? sender, DeviceMovedEventArgs e)
    {
        // 初期化中またはペアリング中はマウス入力を無視
        if (_isInitializing || _isPairing)
            return;
        
        // 接続状況に応じて自動切替の一時停止/再開を判定（機能が有効な場合のみ）
        EvaluateAutoSwitchSuspension();
        
        // Track button press state per device for drag/hold locking
        if (e.ButtonFlags != 0)
        {
            int deltaDown = 0;
            int deltaUp = 0;

            if (e.ButtonFlags.HasFlag(RawMouseButtonFlags.RI_MOUSE_LEFT_BUTTON_DOWN)) deltaDown++;
            if (e.ButtonFlags.HasFlag(RawMouseButtonFlags.RI_MOUSE_RIGHT_BUTTON_DOWN)) deltaDown++;
            if (e.ButtonFlags.HasFlag(RawMouseButtonFlags.RI_MOUSE_MIDDLE_BUTTON_DOWN)) deltaDown++;
            if (e.ButtonFlags.HasFlag(RawMouseButtonFlags.RI_MOUSE_BUTTON_4_DOWN)) deltaDown++;
            if (e.ButtonFlags.HasFlag(RawMouseButtonFlags.RI_MOUSE_BUTTON_5_DOWN)) deltaDown++;

            if (e.ButtonFlags.HasFlag(RawMouseButtonFlags.RI_MOUSE_LEFT_BUTTON_UP)) deltaUp++;
            if (e.ButtonFlags.HasFlag(RawMouseButtonFlags.RI_MOUSE_RIGHT_BUTTON_UP)) deltaUp++;
            if (e.ButtonFlags.HasFlag(RawMouseButtonFlags.RI_MOUSE_MIDDLE_BUTTON_UP)) deltaUp++;
            if (e.ButtonFlags.HasFlag(RawMouseButtonFlags.RI_MOUSE_BUTTON_4_UP)) deltaUp++;
            if (e.ButtonFlags.HasFlag(RawMouseButtonFlags.RI_MOUSE_BUTTON_5_UP)) deltaUp++;

            if (deltaDown != 0 || deltaUp != 0)
            {
                int current = 0;
                _buttonsDownCountByDevice.TryGetValue(e.DeviceId, out current);
                current = Math.Max(0, current + deltaDown - deltaUp);
                _buttonsDownCountByDevice[e.DeviceId] = current;
            }
        }

        // 排他モードが有効な場合のみアクティブマウス制御を適用
        if (_config?.ExclusiveActiveMouse == true)
        {
            var now = DateTime.Now;
            int holdMs = Math.Max(50, Math.Min(1000, _config.ActiveHoldMs));
            bool isSameAsActive = _activeDeviceId == e.DeviceId;
            bool activeButtonsHeld = _activeDeviceId != null &&
                                     _buttonsDownCountByDevice.TryGetValue(_activeDeviceId, out var cnt) && cnt > 0;
            bool activeValidTime = _activeDeviceId != null && now < _activeUntil;
            bool lockToActive = activeButtonsHeld || activeValidTime;

            if (_activeDeviceId == null)
            {
                _activeDeviceId = e.DeviceId;
                _activeUntil = now.AddMilliseconds(holdMs);
            }
            else if (isSameAsActive)
            {
                // 同じアクティブデバイス → 猶予延長
                _activeUntil = now.AddMilliseconds(holdMs);
            }
            else
            {
                // 他デバイスの入力
                if (lockToActive)
                {
                    // ボタンが押下されたまま、または猶予内 → 無視
                    return;
                }
                else
                {
                    // ロックが無い → アクティブ切替
                    _activeDeviceId = e.DeviceId;
                    _activeUntil = now.AddMilliseconds(holdMs);
                }
            }
        }

        // ここに来るのは「アクティブなマウス」からの入力のみ
        _policy?.HandleMovement(e.DeviceId, e.DeltaX, e.DeltaY);

        if (_config?.EnableMousePositionMemory == true)
        {
            _mousePositionManager?.UpdateDevicePosition(e.DeviceId, e.DeltaX, e.DeltaY);
        }
    }

    private void OnDeviceConnectionChanged(object? sender, DeviceConnectionChangedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(e.DeviceId)) return;
        if (e.Arrived)
            _connectedDevicePaths.Add(e.DeviceId);
        else
            _connectedDevicePaths.Remove(e.DeviceId);

        EvaluateAutoSwitchSuspension();
    }

    /// <summary>
    /// 接続状況から「マウスが1台だけ」を判定し、
    /// 自動切り替えの一時停止/再開とバルーン表示を行う
    /// </summary>
    private void EvaluateAutoSwitchSuspension()
    {
        if (_config == null || _policy == null || _tray == null) return;
        // 機能が無効なら何もしない
        if (!_config.EnableDeviceConnectionCheck) return;
        // 設定上のデバイスが2台未満なら対象外
        var configured = new System.Collections.Generic.HashSet<string>(System.StringComparer.OrdinalIgnoreCase);
        foreach (var dev in _config.Devices.Values)
        {
            if (!string.IsNullOrWhiteSpace(dev.DevicePath)) configured.Add(dev.DevicePath);
        }
        if (configured.Count < 2) return;

        // 1) 接続状況での判定を優先
        int connectedConfiguredCount = 0;
        foreach (var path in configured)
        {
            if (_connectedDevicePaths.Contains(path)) connectedConfiguredCount++;
        }

        if (connectedConfiguredCount == 1)
        {
            if (!_autoSwitchTemporarilyDisabled)
            {
                _autoSwitchTemporarilyDisabled = true;
                _policy.AutoSwitchEnabled = false;
                _tray.ShowNotification(
                    "接続中のマウスが1台になりました。自動切り替えを一時停止します。",
                    ToolTipIcon.Warning
                );
            }
        }
        else if (connectedConfiguredCount >= 2)
        {
            if (_autoSwitchTemporarilyDisabled)
            {
                _autoSwitchTemporarilyDisabled = false;
                _policy.AutoSwitchEnabled = true;
                _tray.ShowNotification(
                    "2台のマウスを検出。自動切り替えを再開します。",
                    ToolTipIcon.Info
                );
            }
        }
    }

    private void OnPairingStarted(object? sender, EventArgs e)
    {
        _isPairing = true;
        System.Diagnostics.Debug.WriteLine("[Program] ペアリング開始 - マウス入力処理を無効化");
    }

    private void OnPairingCompleted(object? sender, EventArgs e)
    {
        _isPairing = false;
        System.Diagnostics.Debug.WriteLine("[Program] ペアリング完了 - マウス入力処理を有効化");
    }

    private void OnPairingRequested(object? sender, EventArgs e)
    {
        // ペアリング完了後、仮想デスクトップ枚数を再確保
        _vdController?.EnsureDesktopCount(_config?.VirtualDesktopTargetCount ?? 2);
        
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
