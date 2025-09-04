using System;
using System.Windows.Forms;
using TwoMiceVD.Core;
using TwoMiceVD.Input;
using TwoMiceVD.UI;
using TwoMiceVD.Configuration;

namespace TwoMiceVD;

internal class Program : ApplicationContext
{
    private readonly RawInputManager? _rawInput;
    private readonly VirtualDesktopController? _vdController;
    private readonly SwitchPolicy? _policy;
    private readonly TrayUI? _tray;
    private readonly ConfigStore? _config;
    private readonly MousePositionManager? _mousePositionManager;
    
    // 直近で動いているマウス（アクティブ）の判定用
    private string? _activeDeviceId;
    private DateTime _activeUntil = DateTime.MinValue;
    private int _activeHoldMs = 150; // 連続動作の猶予時間（設定連動）

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
            
            // マウス位置管理の初期化
            _mousePositionManager = new MousePositionManager();

            // アクティブ猶予時間を設定から反映
            _activeHoldMs = Math.Max(50, Math.Min(1000, _config.ActiveHoldMs));

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

    private void OnDeviceMoved(object? sender, DeviceMovedEventArgs e)
    {
        // 排他モードが有効な場合のみアクティブマウス制御を適用
        if (_config?.ExclusiveActiveMouse == true)
        {
            var now = DateTime.Now;
            int holdMs = Math.Max(50, Math.Min(1000, _config.ActiveHoldMs));
            bool activeValid = _activeDeviceId != null && now < _activeUntil;
            bool isSameAsActive = _activeDeviceId == e.DeviceId;

            if (!activeValid)
            {
                // アクティブ期限切れ → 現在のデバイスを新たにアクティブに
                _activeDeviceId = e.DeviceId;
                _activeUntil = now.AddMilliseconds(holdMs);
            }
            else if (!isSameAsActive)
            {
                // 他デバイスの入力は無視（保存も切替も行わない）
                return;
            }
            else
            {
                // 同じアクティブデバイス → 猶予延長
                _activeUntil = now.AddMilliseconds(holdMs);
            }
        }

        // ここに来るのは「アクティブなマウス」からの入力のみ
        _policy?.HandleMovement(e.DeviceId, e.DeltaX, e.DeltaY);

        if (_config?.EnableMousePositionMemory == true)
        {
            _mousePositionManager?.UpdateDevicePosition(e.DeviceId, e.DeltaX, e.DeltaY);
        }
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
