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
        // 先に仮想デスクトップの切り替えを判定・実行
        // （マウス位置の復元を最後にすることで上書きを防ぐ）
        _policy?.HandleMovement(e.DeviceId, e.DeltaX, e.DeltaY);

        // 設定が有効な場合のみマウス位置管理を実行
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
