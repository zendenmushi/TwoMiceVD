using System;
using System.Collections.Generic;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;
using TwoMiceVD.Configuration;
using TwoMiceVD.Interop;
// H.InputSimulator 1.5.0 exposes WindowsInput-compatible namespaces
using WindowsInput;
using static TwoMiceVD.Interop.NativeMethods;

namespace TwoMiceVD.Core;

public class VirtualDesktopController : IDisposable
{
    private int _currentDesktopIndex = 0;
    private readonly InputSimulator _inputSimulator;
    private readonly Dictionary<int, Point> _desktopCursorPositions = new Dictionary<int, Point>();
    private VirtualDesktopManagerWrapper? _vdManager;
    private FastDesktopDetector? _fastDetector;
    
    /// <summary>
    /// 最後に切り替えたデスクトップIDを記憶（オプション無効時の比較用）
    /// </summary>
    private string? _lastSwitchedDesktopId = null;
    private bool _disposed = false;
    private readonly ConfigStore? _config;
    private readonly Control _uiInvoker;

    public VirtualDesktopController(ConfigStore? config)
    {
        _config = config;
        // UIスレッドに紐づくInvokerを作成（このコンストラクタはUIスレッド上で呼ばれる想定）
        _uiInvoker = new Control();
        _uiInvoker.CreateControl();
        
        // InputSimulatorのインスタンスを初期化
        _inputSimulator = new InputSimulator();
        
        // COMインターフェースの初期化を試行
        try
        {
            _vdManager = new VirtualDesktopManagerWrapper();
            
            // FastDesktopDetectorを初期化
            if (_config != null)
            {
                _fastDetector = new FastDesktopDetector(_vdManager, _config, this, _uiInvoker);
            }
        }
        catch
        {
            // COM初期化失敗時はSendInputのみで動作
            _vdManager = null;
            _fastDetector = null;
        }
    }

    public void EnsureDesktopCount(int target)
    {
        // 仮想デスクトップ2枚の存在を前提とするため、何も行わない
        // ユーザーが事前にWin+Ctrl+Dで仮想デスクトップを2枚作成済みと想定
    }

    public void SwitchTo(string desktopId)
    {
        System.Diagnostics.Debug.WriteLine($"[TwoMiceVD] SwitchTo呼び出し: desktopId={desktopId}");

        // 現在のデスクトップを確認するかどうか
        if (_config?.CheckCurrentDesktop ?? true)
        {
            // オプション有効: 実際の現在デスクトップを確認
            if (_fastDetector != null && _fastDetector.TryGetCurrentDesktopId(out string? currentDesktopId) && currentDesktopId != null)
            {
                // 高速判定成功：GUIDをデスクトップIDに変換して比較
                string? targetDesktopId = ConvertToDesktopId(desktopId);
                if (targetDesktopId != null && currentDesktopId == targetDesktopId)
                {
                    System.Diagnostics.Debug.WriteLine($"[TwoMiceVD] 既に目的のデスクトップ({targetDesktopId})にいます - 切り替えをスキップ");
                    _lastSwitchedDesktopId = targetDesktopId;  // 記憶を更新
                    return;
                }
                System.Diagnostics.Debug.WriteLine($"[TwoMiceVD] デスクトップ切り替えが必要: {currentDesktopId} → {targetDesktopId}");
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("[TwoMiceVD] 高速判定失敗、通常の切り替え処理を実行");
            }
        }
        else
        {
            // オプション無効: 記憶している位置と比較
            string? targetDesktopId = ConvertToDesktopId(desktopId);
            if (targetDesktopId != null && _lastSwitchedDesktopId == targetDesktopId)
            {
                System.Diagnostics.Debug.WriteLine($"[TwoMiceVD] 記憶上、既に{targetDesktopId}にいるはず - 切り替えをスキップ");
                return;
            }
            System.Diagnostics.Debug.WriteLine($"[TwoMiceVD] 記憶に基づく切り替え: {_lastSwitchedDesktopId ?? "不明"} → {targetDesktopId}");
        }
        
        // GUIDかどうかを判定
        if (Guid.TryParse(desktopId, out Guid targetGuid))
        {
            // GUIDモード: 実際はGUIDからインデックスを推定してSendInputで切り替え
            System.Diagnostics.Debug.WriteLine($"[TwoMiceVD] GUIDモード: {targetGuid}");
            
            // ConfigStoreからGUIDとインデックスのマッピングを汎用に取得
            if (_config != null)
            {
                string? targetKey = _config.TryGetKeyByGuid(targetGuid);
                if (targetKey != null)
                {
                    System.Diagnostics.Debug.WriteLine($"[TwoMiceVD] GUID→デスクトップキー変換: {targetGuid} → {targetKey}");
                    SwitchUsingShortcuts(targetKey);
                    // 切り替え成功後、記憶を更新
                    _lastSwitchedDesktopId = ConvertToDesktopId(desktopId);
                    return;
                }
                
                System.Diagnostics.Debug.WriteLine($"[TwoMiceVD] GUID→デスクトップキー変換失敗、VD0として処理: {targetGuid}");
            }
            
            // GUID解決不可のためショートカット方式で切り替え（VD0と仮定）
            SwitchUsingShortcuts("VD0");
            _lastSwitchedDesktopId = ConvertToDesktopId("VD0");
        }
        else
        {
            // インデックスモード: ショートカット方式で切り替え
            SwitchUsingShortcuts(desktopId);
            // 切り替え成功後、記憶を更新
            _lastSwitchedDesktopId = ConvertToDesktopId(desktopId);
        }
    }

    private void SwitchUsingShortcuts(string desktopId)
    {
        // "VD{n}" からインデックス n を取得（将来の多枚数対応）
        int targetIndex;
        if (_config != null && _config.TryGetIndexByKey(desktopId) is int idx)
        {
            targetIndex = idx;
        }
        else if (int.TryParse(desktopId.Replace("VD", ""), out int parsed))
        {
            targetIndex = parsed;
        }
        else
        {
            System.Diagnostics.Debug.WriteLine($"[TwoMiceVD] 不正なdesktopId: {desktopId}");
            return;
        }

        // 手動切替などでズレた場合に備えて、現在のインデックスを高速判定で同期
        bool synced = SyncCurrentDesktopIndex();

        // 現在地と目標が一致していればスキップ
        if (synced && targetIndex == _currentDesktopIndex)
        {
            System.Diagnostics.Debug.WriteLine("[TwoMiceVD] 同期後、既に目的インデックスのためスキップ");
            return;
        }

        // 目標インデックスへ移動（将来の多枚数対応：複数回送出）
        int delta = targetIndex - _currentDesktopIndex;
        if (delta > 0)
        {
            for (int i = 0; i < delta; i++) SendDesktopRightShortcut();
        }
        else if (delta < 0)
        {
            for (int i = 0; i < -delta; i++) SendDesktopLeftShortcut();
        }

        _currentDesktopIndex = targetIndex;
    }

    /// <summary>
    /// FastDetector が取得した現在のデスクトップIDから _currentDesktopIndex を同期する
    /// 戻り値: 同期に成功した場合 true（VD0/VD1 いずれかを検出できた）
    /// </summary>
    private bool SyncCurrentDesktopIndex()
    {
        try
        {
            string? detected = GetCurrentDesktopIdFast();
            if (detected == "VD0")
            {
                _currentDesktopIndex = 0;
                return true;
            }
            if (detected == "VD1")
            {
                _currentDesktopIndex = 1;
                return true;
            }
        }
        catch { /* 無視して既存値を維持 */ }
        return false;
    }

    public void SendDesktopLeftShortcut()
    {
        // Win+Ctrl+左矢印でデスクトップ1へ切り替え
        _inputSimulator.Keyboard.ModifiedKeyStroke(
            new[] { VirtualKeyCode.LWIN, VirtualKeyCode.CONTROL },
            VirtualKeyCode.LEFT
        );
    }

    public void SendDesktopRightShortcut()
    {
        // Win+Ctrl+右矢印でデスクトップ2へ切り替え
        _inputSimulator.Keyboard.ModifiedKeyStroke(
            new[] { VirtualKeyCode.LWIN, VirtualKeyCode.CONTROL },
            VirtualKeyCode.RIGHT
        );
    }

    /// <summary>
    /// 新しい仮想デスクトップを作成してGUIDを取得する
    /// </summary>
    public async Task<(Guid vd0, Guid vd1)?> CreateDesktopAsync()
    {
        System.Diagnostics.Debug.WriteLine("[TwoMiceVD] 新しい仮想デスクトップを作成中...");
        
        // メインデスクトップのGUIDを先に取得
        var vd0Guid = await GetCurrentDesktopGuidAsync();
        if (vd0Guid == null)
        {
            System.Diagnostics.Debug.WriteLine("[TwoMiceVD] メインデスクトップのGUID取得に失敗");
            return null;
        }
        
        // Win+Ctrl+D で新しいデスクトップを作成
        _inputSimulator.Keyboard.ModifiedKeyStroke(
            new[] { VirtualKeyCode.LWIN, VirtualKeyCode.CONTROL },
            VirtualKeyCode.VK_D
        );
        
        // 作成完了を待機（アニメーションとデスクトップ初期化）
        await Task.Delay(800);
        
        // 新しく作成されたデスクトップのGUIDを取得
        var vd1Guid = await GetCurrentDesktopGuidAsync();
        if (vd1Guid == null)
        {
            System.Diagnostics.Debug.WriteLine("[TwoMiceVD] 新しいデスクトップのGUID取得に失敗");
            // メインに戻る
            SendDesktopLeftShortcut();
            await Task.Delay(250);
            return null;
        }
        
        // メイン（左端）デスクトップに戻る
        System.Diagnostics.Debug.WriteLine("[TwoMiceVD] メインデスクトップに戻る...");
        SendDesktopLeftShortcut();
        
        // 移動アニメーション完了を待機
        await Task.Delay(250);
        
        System.Diagnostics.Debug.WriteLine($"[TwoMiceVD] 仮想デスクトップ作成完了: VD0={vd0Guid}, VD1={vd1Guid}");
        return (vd0Guid.Value, vd1Guid.Value);
    }

    /// <summary>
    /// 仮想デスクトップのGUIDを簡易的に取得する（検証なし）
    /// </summary>
    public async Task<(Guid vd0, Guid vd1)?> GetDesktopGuidsAsync()
    {
        if (_vdManager?.IsAvailable != true)
            return null;

        System.Diagnostics.Debug.WriteLine("[TwoMiceVD] 仮想デスクトップGUIDを取得中...");

        // メインデスクトップのGUIDを取得
        var vd0Guid = await GetCurrentDesktopGuidAsync();
        if (vd0Guid == null)
        {
            System.Diagnostics.Debug.WriteLine("[TwoMiceVD] VD0 GUIDの取得に失敗");
            return null;
        }

        // 右のデスクトップに移動してGUID取得
        SendDesktopRightShortcut();
        await Task.Delay(250);
        
        var vd1Guid = await GetCurrentDesktopGuidAsync();
        if (vd1Guid == null)
        {
            System.Diagnostics.Debug.WriteLine("[TwoMiceVD] VD1 GUIDの取得に失敗");
            
            // メインに戻してからnullを返す
            SendDesktopLeftShortcut();
            await Task.Delay(250);
            return null;
        }

        // メインに戻る
        SendDesktopLeftShortcut();
        await Task.Delay(250);

        System.Diagnostics.Debug.WriteLine($"[TwoMiceVD] GUID取得完了: VD0={vd0Guid}, VD1={vd1Guid}");
        return (vd0Guid.Value, vd1Guid.Value);
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
    private void RestoreCursorPosition(int desktopIndex)
    {
        if (_desktopCursorPositions.TryGetValue(desktopIndex, out Point position))
        {
            // Win32 APIのSetCursorPosを使用してマウス位置を設定
            SetCursorPos(position.X, position.Y);
        }
    }

    #region Probe Window and GUID Operations

    /// <summary>
    /// 不可視のProbeウィンドウを作成してデスクトップGUIDを取得
    /// </summary>
    public async Task<Guid?> GetCurrentDesktopGuidAsync()
    {
        System.Diagnostics.Debug.WriteLine("[VDController] GetCurrentDesktopGuidAsync開始");
        
        if (_vdManager?.IsAvailable != true)
        {
            System.Diagnostics.Debug.WriteLine("[VDController] VirtualDesktopManagerが利用できません");
            return null;
        }

        Form? probeWindow = null;
        try
        {
            System.Diagnostics.Debug.WriteLine("[VDController] プローブウィンドウを作成中...");
            probeWindow = CreateProbeWindow();
            
            if (probeWindow == null)
            {
                System.Diagnostics.Debug.WriteLine("[VDController] プローブウィンドウの作成に失敗");
                return null;
            }
            var hwnd = GetFormHandleThreadSafe(probeWindow);
            if (hwnd == IntPtr.Zero)
            {
                System.Diagnostics.Debug.WriteLine("[VDController] プローブウィンドウのハンドルが無効");
                return null;
            }
            System.Diagnostics.Debug.WriteLine($"[VDController] プローブウィンドウ作成完了: Handle=0x{hwnd:X}");
            
            // ウィンドウが確実に作成されるまで待機
            await Task.Delay(300); // 待機時間をさらに延長
            
            System.Diagnostics.Debug.WriteLine("[VDController] デスクトップGUIDを取得中...");
            var guid = _vdManager.GetWindowDesktopId(hwnd);
            
            if (guid == null)
            {
                System.Diagnostics.Debug.WriteLine("[VDController] GUID取得失敗 - プローブウィンドウが適切なデスクトップに配置されていない可能性");
            }
            else
            {
                System.Diagnostics.Debug.WriteLine($"[VDController] GUID取得成功: {guid}");
            }
            
            return guid;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[VDController] GetCurrentDesktopGuidAsync中にエラー: {ex.Message}");
            return null;
        }
        finally
        {
            if (probeWindow != null)
            {
                System.Diagnostics.Debug.WriteLine("[VDController] プローブウィンドウを破棄");
                DisposeFormThreadSafe(probeWindow);
            }
        }
    }

    /// <summary>
    /// デスクトップ構成の検証（メイン判定と枚数チェック）
    /// </summary>
    public async Task<DesktopValidationResult> ValidateDesktopConfigurationAsync()
    {
        if (_vdManager?.IsAvailable != true)
        {
            return new DesktopValidationResult
            {
                IsValid = false,
                ErrorType = DesktopValidationErrorType.ComInterfaceUnavailable,
                ErrorMessage = "仮想デスクトップCOMインターフェースが利用できません。SendInputモードで動作します。"
            };
        }

        try
        {
            // 1. 現在のデスクトップGUIDを取得
            var baseGuid = await GetCurrentDesktopGuidAsync();
            if (baseGuid == null)
            {
                return new DesktopValidationResult
                {
                    IsValid = false,
                    ErrorType = DesktopValidationErrorType.GuidRetrievalFailed,
                    ErrorMessage = "現在のデスクトップGUIDを取得できませんでした。"
                };
            }

            // 2. メイン（左端）判定
            SendDesktopLeftShortcut();
            await Task.Delay(250); // アニメーション待機

            var leftGuid = await GetCurrentDesktopGuidAsync();
            if (leftGuid != baseGuid)
            {
                // 元の位置に戻す試行
                SendDesktopRightShortcut();
                await Task.Delay(250);
                
                return new DesktopValidationResult
                {
                    IsValid = false,
                    ErrorType = DesktopValidationErrorType.NotOnMainDesktop,
                    ErrorMessage = "ペアリングはメイン（左端）のデスクトップで開始してください。メインへ切り替えてから再度お試しください。"
                };
            }

            // 3. 枚数（右側）判定
            SendDesktopRightShortcut();
            await Task.Delay(250);

            var right1Guid = await GetCurrentDesktopGuidAsync();
            if (right1Guid == baseGuid)
            {
                return new DesktopValidationResult
                {
                    IsValid = false,
                    ErrorType = DesktopValidationErrorType.NotEnoughDesktops,
                    ErrorMessage = "仮想デスクトップが1枚です。新しいデスクトップを作成してから再度ペアリングしてください。"
                };
            }

            // 4. 3枚以上の確認
            SendDesktopRightShortcut();
            await Task.Delay(250);

            var right2Guid = await GetCurrentDesktopGuidAsync();
            if (right2Guid != right1Guid && right2Guid != baseGuid)
            {
                // メインへ戻す
                await ReturnToDesktopAsync(baseGuid.Value);
                
                return new DesktopValidationResult
                {
                    IsValid = false,
                    ErrorType = DesktopValidationErrorType.TooManyDesktops,
                    ErrorMessage = "仮想デスクトップが3枚以上あります。本アプリは2枚までを想定しています。メインへ戻して構成を見直してから再度お試しください。"
                };
            }

            // 5. 成功時：メインに戻してGUIDを返す
            await ReturnToDesktopAsync(baseGuid.Value);
            
            return new DesktopValidationResult
            {
                IsValid = true,
                ErrorType = DesktopValidationErrorType.None,
                VD0Guid = baseGuid.Value,
                VD1Guid = right1Guid.Value
            };
        }
        catch (Exception ex)
        {
            return new DesktopValidationResult
            {
                IsValid = false,
                ErrorType = DesktopValidationErrorType.GuidRetrievalFailed,
                ErrorMessage = $"デスクトップ構成の検証中にエラーが発生しました: {ex.Message}"
            };
        }
    }

    /// <summary>
    /// 指定されたGUIDのデスクトップに復帰を試行
    /// </summary>
    public async Task<bool> ReturnToDesktopAsync(Guid targetGuid)
    {
        if (_vdManager?.IsAvailable != true)
            return false;

        // 最大5回試行で指定デスクトップへ移動
        for (int attempt = 0; attempt < 5; attempt++)
        {
            var currentGuid = await GetCurrentDesktopGuidAsync();
            if (currentGuid == targetGuid)
                return true;

            // 左右に移動を試行
            if (attempt % 2 == 0)
                SendDesktopLeftShortcut();
            else
                SendDesktopRightShortcut();

            await Task.Delay(200);
        }

        return false;
    }

    /// <summary>
    /// 保存されたGUIDの存在を検証
    /// </summary>
    public async Task<bool> ValidateStoredGuidsAsync(Guid vd0Guid, Guid vd1Guid)
    {
        if (_vdManager?.IsAvailable != true)
            return false;

        Form? probeWindow = null;
        try
        {
            probeWindow = CreateProbeWindow();
            await Task.Delay(50);

            // 両方のGUIDが有効かテスト
            var hwnd = probeWindow != null ? GetFormHandleThreadSafe(probeWindow) : IntPtr.Zero;
            if (hwnd == IntPtr.Zero) return false;
            bool vd0Valid = _vdManager.MoveWindowToDesktop(hwnd, vd0Guid);
            bool vd1Valid = _vdManager.MoveWindowToDesktop(hwnd, vd1Guid);

            if (!vd0Valid)
            {
                System.Diagnostics.Debug.WriteLine($"[TwoMiceVD] GUID検証: VD0 GUID({vd0Guid})が無効です");
            }
            
            if (!vd1Valid)
            {
                System.Diagnostics.Debug.WriteLine($"[TwoMiceVD] GUID検証: VD1 GUID({vd1Guid})が無効です");
            }

            return vd0Valid && vd1Valid;
        }
        finally
        {
            if (probeWindow != null) DisposeFormThreadSafe(probeWindow);
        }
    }

    private Form CreateProbeWindow()
    {
        Form CreateOnUi()
        {
            var form = new Form
            {
                FormBorderStyle = FormBorderStyle.None,
                ShowInTaskbar = false,
                WindowState = FormWindowState.Normal,
                Size = new Size(1, 1),
                StartPosition = FormStartPosition.Manual,
                Location = new Point(-10000, -10000),
                Visible = true
            };

            // 確実にウィンドウハンドルを作成
            var handle = form.Handle;
            Application.DoEvents();
            form.Visible = false;
            System.Diagnostics.Debug.WriteLine($"[TwoMiceVD] Probeウィンドウ作成: HWND=0x{handle:X}");
            return form;
        }

        if (_uiInvoker.InvokeRequired)
        {
            Form? created = null;
            _uiInvoker.Invoke(new Action(() => created = CreateOnUi()));
            return created!;
        }
        else
        {
            return CreateOnUi();
        }
    }

    #endregion

    private static IntPtr GetFormHandleThreadSafe(Form form)
    {
        if (form.IsDisposed) return IntPtr.Zero;
        if (form.InvokeRequired)
        {
            IntPtr handle = IntPtr.Zero;
            try
            {
                form.Invoke(new Action(() => { if (!form.IsDisposed) handle = form.Handle; }));
            }
            catch { return IntPtr.Zero; }
            return handle;
        }
        return form.Handle;
    }

    private static void DisposeFormThreadSafe(Form form)
    {
        if (form.IsDisposed) return;
        if (form.InvokeRequired)
        {
            try { form.Invoke(new Action(() => { if (!form.IsDisposed) form.Dispose(); })); }
            catch { /* ignore */ }
        }
        else
        {
            try { form.Dispose(); } catch { }
        }
    }

    #region FastDesktopDetector Methods

    /// <summary>
    /// マーカーシステムを初期化します
    /// </summary>
    public async Task<bool> InitializeMarkersAsync()
    {
        if (_fastDetector != null)
        {
            return await _fastDetector.InitializeMarkersAsync();
        }
        return false;
    }

    /// <summary>
    /// 現在のデスクトップIDを高速に取得します
    /// </summary>
    public string? GetCurrentDesktopIdFast()
    {
        if (_fastDetector != null)
        {
            return _fastDetector.GetCurrentDesktopId();
        }
        return null;
    }

    /// <summary>
    /// 現在のデスクトップIDを取得します（自己修復機能付き）
    /// </summary>
    public async Task<string?> GetCurrentDesktopIdWithRecoveryAsync()
    {
        if (_fastDetector != null)
        {
            return await _fastDetector.GetCurrentDesktopIdWithRecovery();
        }
        return null;
    }

    /// <summary>
    /// 全マーカーの有効性を検証し、必要に応じて修復します
    /// </summary>
    public bool ValidateAndRepairMarkers()
    {
        if (_fastDetector != null)
        {
            return _fastDetector.ValidateAndRepairAll();
        }
        return false;
    }

    /// <summary>
    /// マーカーシステムが初期化されているかを確認します
    /// </summary>
    public bool IsMarkersInitialized => _fastDetector?.IsInitialized ?? false;

    /// <summary>
    /// マーカーシステムの前提条件をチェックします
    /// </summary>
    public async Task<(bool canInitialize, string reason)> CheckMarkerInitializationPrerequisitesAsync()
    {
        if (_fastDetector != null)
        {
            return await _fastDetector.CheckInitializationPrerequisitesAsync();
        }
        return (false, "FastDetectorが初期化されていません");
    }

    #endregion

    #region Helper Methods

    /// <summary>
    /// GUID文字列またはデスクトップIDをデスクトップID（VD0/VD1）に変換
    /// </summary>
    private string? ConvertToDesktopId(string input)
    {
        // 既にデスクトップIDの場合はそのまま返す
        if (input == "VD0" || input == "VD1")
            return input;

        // GUIDの場合は設定から対応するデスクトップIDを検索
        if (Guid.TryParse(input, out Guid targetGuid) && _config != null)
        {
            if (_config.VirtualDesktops.Ids.TryGetValue("VD0", out string? vd0Guid) && 
                vd0Guid == targetGuid.ToString())
            {
                return "VD0";
            }
            else if (_config.VirtualDesktops.Ids.TryGetValue("VD1", out string? vd1Guid) && 
                     vd1Guid == targetGuid.ToString())
            {
                return "VD1";
            }
        }

        return null;
    }

    #endregion

    public void Dispose()
    {
        if (!_disposed)
        {
            _fastDetector?.Dispose();
            _vdManager?.Dispose();
            _disposed = true;
        }
    }

    public enum DesktopValidationErrorType
    {
        None,                  // エラーなし（成功）
        NotEnoughDesktops,     // デスクトップが1枚（不足）
        TooManyDesktops,       // デスクトップが3枚以上（過多）
        NotOnMainDesktop,      // メイン（左端）デスクトップにいない
        ComInterfaceUnavailable, // COMインターフェースが利用できない
        GuidRetrievalFailed    // GUID取得に失敗
    }

    public class DesktopValidationResult
    {
        public bool IsValid { get; set; }
        public DesktopValidationErrorType ErrorType { get; set; } = DesktopValidationErrorType.None;
        public string? ErrorMessage { get; set; }
        public Guid VD0Guid { get; set; }
        public Guid VD1Guid { get; set; }
    }
}
