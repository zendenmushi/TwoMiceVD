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
    private bool _disposed = false;
    private readonly ConfigStore? _config;

    public VirtualDesktopController(ConfigStore? config)
    {
        _config = config;
        
        // InputSimulatorのインスタンスを初期化
        _inputSimulator = new InputSimulator();
        
        // COMインターフェースの初期化を試行
        try
        {
            _vdManager = new VirtualDesktopManagerWrapper();
        }
        catch
        {
            // COM初期化失敗時はSendInputのみで動作
            _vdManager = null;
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
        
        // GUIDかどうかを判定
        if (Guid.TryParse(desktopId, out Guid targetGuid))
        {
            // GUIDモード: 実際はGUIDからインデックスを推定してSendInputで切り替え
            System.Diagnostics.Debug.WriteLine($"[TwoMiceVD] GUIDモード: {targetGuid}");
            
            // ConfigStoreからGUIDとインデックスのマッピングを取得
            if (_config != null)
            {
                string? targetIndexStr = null;
                
                // VD0とVD1のGUIDを確認
                if (_config.VirtualDesktops.Ids.TryGetValue("VD0", out string? vd0Guid) && vd0Guid == targetGuid.ToString())
                {
                    targetIndexStr = "VD0";
                }
                else if (_config.VirtualDesktops.Ids.TryGetValue("VD1", out string? vd1Guid) && vd1Guid == targetGuid.ToString())
                {
                    targetIndexStr = "VD1";
                }
                
                if (targetIndexStr != null)
                {
                    System.Diagnostics.Debug.WriteLine($"[TwoMiceVD] GUID→インデックス変換: {targetGuid} → {targetIndexStr}");
                    FallbackToSendInput(targetIndexStr);
                    return;
                }
                
                System.Diagnostics.Debug.WriteLine($"[TwoMiceVD] GUID→インデックス変換失敗、VD0として処理: {targetGuid}");
            }
            
            // フォールバックとしてSendInputを使用（VD0と仮定）
            FallbackToSendInput("VD0");
        }
        else
        {
            // インデックスモード: 従来の処理
            FallbackToSendInput(desktopId);
        }
    }

    private void FallbackToSendInput(string desktopId)
    {
        // "VD0" または "VD1" からインデックス 0 または 1 を取得
        if (int.TryParse(desktopId.Replace("VD", ""), out int targetIndex))
        {
            System.Diagnostics.Debug.WriteLine($"[TwoMiceVD] SendInputモードで切り替え: インデックス={targetIndex}");
            
            if (targetIndex == _currentDesktopIndex) return;

            if (targetIndex == 0)
            {
                SendDesktopLeftShortcut(); // デスクトップ1へ切り替え
            }
            else if (targetIndex == 1)
            {
                SendDesktopRightShortcut(); // デスクトップ2へ切り替え  
            }

            _currentDesktopIndex = targetIndex;
        }
        else
        {
            System.Diagnostics.Debug.WriteLine($"[TwoMiceVD] 不正なdesktopId: {desktopId}");
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
        if (_vdManager?.IsAvailable != true)
            return null;

        Form? probeWindow = null;
        try
        {
            probeWindow = CreateProbeWindow();
            
            // ウィンドウが確実に作成されるまで待機
            await Task.Delay(200); // 待機時間を延長
            
            var guid = _vdManager.GetWindowDesktopId(probeWindow.Handle);
            if (guid == null)
            {
                System.Diagnostics.Debug.WriteLine($"[TwoMiceVD] GetCurrentDesktopGuid: GUID取得失敗 (TYPE_E_ELEMENTNOTFOUND の可能性)");
            }
            else
            {
                System.Diagnostics.Debug.WriteLine($"[TwoMiceVD] GetCurrentDesktopGuid: 取得成功 GUID={guid}");
            }
            
            return guid;
        }
        finally
        {
            probeWindow?.Dispose();
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
                    ErrorMessage = "仮想デスクトップが3枚以上あります。本アプリは2枚までを想定しています。メインへ戻して構成を見直してから再度お試しください。"
                };
            }

            // 5. 成功時：メインに戻してGUIDを返す
            await ReturnToDesktopAsync(baseGuid.Value);
            
            return new DesktopValidationResult
            {
                IsValid = true,
                VD0Guid = baseGuid.Value,
                VD1Guid = right1Guid.Value
            };
        }
        catch (Exception ex)
        {
            return new DesktopValidationResult
            {
                IsValid = false,
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
            bool vd0Valid = _vdManager.MoveWindowToDesktop(probeWindow.Handle, vd0Guid);
            bool vd1Valid = _vdManager.MoveWindowToDesktop(probeWindow.Handle, vd1Guid);

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
            probeWindow?.Dispose();
        }
    }

    private Form CreateProbeWindow()
    {
        var form = new Form
        {
            FormBorderStyle = FormBorderStyle.None,
            ShowInTaskbar = false,
            WindowState = FormWindowState.Normal, // Minimizedから変更
            Size = new Size(1, 1),
            StartPosition = FormStartPosition.Manual,
            Location = new Point(-10000, -10000),
            Visible = true // falseから変更 - 一時的に表示してからすぐに非表示にする
        };

        // 確実にウィンドウハンドルを作成
        var handle = form.Handle;
        
        // ウィンドウが確実に作成されるまで少し待機
        Application.DoEvents();
        
        // すぐに非表示にする
        form.Visible = false;
        
        System.Diagnostics.Debug.WriteLine($"[TwoMiceVD] Probeウィンドウ作成: HWND=0x{handle:X}");
        
        return form;
    }

    #endregion

    public void Dispose()
    {
        if (!_disposed)
        {
            _vdManager?.Dispose();
            _disposed = true;
        }
    }

    public class DesktopValidationResult
    {
        public bool IsValid { get; set; }
        public string? ErrorMessage { get; set; }
        public Guid VD0Guid { get; set; }
        public Guid VD1Guid { get; set; }
    }
}
