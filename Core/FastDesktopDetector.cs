using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Windows.Forms;
using TwoMiceVD.Configuration;
using TwoMiceVD.Interop;

namespace TwoMiceVD.Core;

public class FastDesktopDetector : IDisposable
{
    private Dictionary<string, DesktopMarkerWindow> _markers;
    private readonly VirtualDesktopManagerWrapper _vdManager;
    private readonly ConfigStore _config;
    private readonly VirtualDesktopController _vdController;
    private readonly Control _uiInvoker;
    private bool _disposed = false;

    public FastDesktopDetector(VirtualDesktopManagerWrapper vdManager, ConfigStore config, VirtualDesktopController vdController, Control uiInvoker)
    {
        _vdManager = vdManager;
        _config = config;
        _vdController = vdController;
        _uiInvoker = uiInvoker;
        _markers = new Dictionary<string, DesktopMarkerWindow>();
    }

    public bool IsInitialized => _markers.Count > 0;

    public async Task<bool> InitializeMarkersAsync()
    {
        try
        {
            Debug.WriteLine("[FastDetector] マーカーシステムを初期化中...");

            // 既存のマーカーをクリーンアップ
            DisposeMarkers();
            Debug.WriteLine("[FastDetector] 既存マーカーのクリーンアップ完了");

            // 設定からGUIDを取得
            Debug.WriteLine($"[FastDetector] 設定確認: Mode={_config.VirtualDesktops.Mode}");
            
            if (_config.VirtualDesktops.Mode != "GUID")
            {
                Debug.WriteLine("[FastDetector] GUIDモードではありません");
                return false;
            }

            if (!_config.VirtualDesktops.Ids.TryGetValue("VD0", out string? vd0String))
            {
                Debug.WriteLine("[FastDetector] VD0のGUIDが設定されていません");
                return false;
            }

            if (!_config.VirtualDesktops.Ids.TryGetValue("VD1", out string? vd1String))
            {
                Debug.WriteLine("[FastDetector] VD1のGUIDが設定されていません");
                return false;
            }

            if (!Guid.TryParse(vd0String, out Guid vd0Guid))
            {
                Debug.WriteLine($"[FastDetector] VD0のGUID形式が無効: {vd0String}");
                return false;
            }

            if (!Guid.TryParse(vd1String, out Guid vd1Guid))
            {
                Debug.WriteLine($"[FastDetector] VD1のGUID形式が無効: {vd1String}");
                return false;
            }

            Debug.WriteLine($"[FastDetector] GUID取得完了 - VD0: {vd0Guid}, VD1: {vd1Guid}");

            // 現在のデスクトップを取得
            Debug.WriteLine("[FastDetector] 現在のデスクトップGUIDを取得中...");
            var currentGuid = await _vdController.GetCurrentDesktopGuidAsync();
            if (currentGuid == null)
            {
                Debug.WriteLine("[FastDetector] 現在のデスクトップGUIDを取得できませんでした");
                return false;
            }

            Debug.WriteLine($"[FastDetector] 現在のデスクトップGUID: {currentGuid.Value}");

            // 現在のデスクトップがVD0かVD1かを判定
            bool isCurrentVD0 = (currentGuid.Value == vd0Guid);
            bool isCurrentVD1 = (currentGuid.Value == vd1Guid);

            Debug.WriteLine($"[FastDetector] デスクトップ判定: isVD0={isCurrentVD0}, isVD1={isCurrentVD1}");

            if (!isCurrentVD0 && !isCurrentVD1)
            {
                Debug.WriteLine($"[FastDetector] 現在のデスクトップ({currentGuid})は設定されたVD0/VD1ではありません");
                Debug.WriteLine($"[FastDetector] 期待値: VD0={vd0Guid}, VD1={vd1Guid}");
                return false;
            }

            // 両方のマーカーを現在のデスクトップで作成し、COM APIで適切なデスクトップに移動
            Debug.WriteLine($"[FastDetector] VD0マーカーを作成中... GUID={vd0Guid}");
            var vd0Marker = CreateMarkerOnUiThread("VD0", vd0Guid);
            _markers["VD0"] = vd0Marker;

            Debug.WriteLine($"[FastDetector] VD1マーカーを作成中... GUID={vd1Guid}");
            var vd1Marker = CreateMarkerOnUiThread("VD1", vd1Guid);
            _markers["VD1"] = vd1Marker;

            // 少し待機してマーカーウィンドウの配置が完了するのを待つ
            //await Task.Delay(200);

            Debug.WriteLine("[FastDetector] マーカーシステムの初期化が完了しました");
            return true;
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[FastDetector] マーカー初期化中にエラー: {ex.Message}");
            Debug.WriteLine($"[FastDetector] エラー詳細: {ex}");
            DisposeMarkers();
            return false;
        }
    }

    /// <summary>
    /// マーカーシステムの前提条件をチェック
    /// </summary>
    public async Task<(bool canInitialize, string reason)> CheckInitializationPrerequisitesAsync()
    {
        try
        {
            // VirtualDesktopManager が利用可能か
            if (_vdManager?.IsAvailable != true)
            {
                return (false, "VirtualDesktopManager が利用できません");
            }

            // GUID モードか
            if (_config.VirtualDesktops.Mode != "GUID")
            {
                return (false, $"GUIDモードではありません (現在: {_config.VirtualDesktops.Mode})");
            }

            // VD0 GUID が存在するか
            if (!_config.VirtualDesktops.Ids.TryGetValue("VD0", out string? vd0String) ||
                !Guid.TryParse(vd0String, out Guid vd0Guid))
            {
                return (false, "VD0のGUIDが無効です");
            }

            // VD1 GUID が存在するか
            if (!_config.VirtualDesktops.Ids.TryGetValue("VD1", out string? vd1String) ||
                !Guid.TryParse(vd1String, out Guid vd1Guid))
            {
                return (false, "VD1のGUIDが無効です");
            }

            // 現在のデスクトップGUIDを取得できるか
            var currentGuid = await _vdController.GetCurrentDesktopGuidAsync();
            if (currentGuid == null)
            {
                return (false, "現在のデスクトップGUIDを取得できません");
            }

            // 現在のデスクトップがVD0またはVD1か
            if (currentGuid.Value != vd0Guid && currentGuid.Value != vd1Guid)
            {
                return (false, $"現在のデスクトップ({currentGuid})がVD0/VD1のどちらでもありません");
            }

            return (true, "すべての前提条件を満たしています");
        }
        catch (Exception ex)
        {
            return (false, $"前提条件チェック中にエラー: {ex.Message}");
        }
    }

    public string? GetCurrentDesktopId()
    {
        if (_disposed || _markers.Count == 0)
            return null;

        try
        {
            // 各マーカーの有効性を確認
            foreach (var kvp in _markers)
            {
                var marker = kvp.Value;
                
                // まずGUID検証を行う
                if (!marker.ValidatePosition())
                {
                    Debug.WriteLine($"[FastDetector] {kvp.Key} のマーカー位置が無効です");
                    continue;
                }

                // 現在のデスクトップにあるかチェック
                if (marker.IsOnCurrentDesktop())
                {
                    Debug.WriteLine($"[FastDetector] 現在のデスクトップ: {kvp.Key}");
                    return kvp.Key;
                }
            }

            Debug.WriteLine("[FastDetector] 現在のデスクトップを特定できませんでした");
            return null;
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[FastDetector] デスクトップ判定中にエラー: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// 現在のデスクトップIDを安全に取得します。失敗時は例外を発生させません。
    /// </summary>
    /// <param name="desktopId">取得されたデスクトップID</param>
    /// <returns>取得に成功した場合はtrue、失敗した場合はfalse</returns>
    public bool TryGetCurrentDesktopId(out string? desktopId)
    {
        desktopId = null;
        try
        {
            desktopId = GetCurrentDesktopId();
            return desktopId != null;
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[FastDetector] 検出失敗: {ex.Message}");
            return false;
        }
    }

    public async Task<string?> GetCurrentDesktopIdWithRecovery()
    {
        try
        {
            // Step 1: 高速判定を試みる
            var quickResult = GetCurrentDesktopId();
            if (quickResult != null)
                return quickResult;

            Debug.WriteLine("[FastDetector] 高速判定に失敗、リカバリーを開始...");

            // Step 2: 現在のGUIDを取得
            var currentGuid = await _vdController.GetCurrentDesktopGuidAsync();
            if (currentGuid == null)
            {
                Debug.WriteLine("[FastDetector] 現在のデスクトップGUIDを取得できませんでした");
                return null;
            }

            // Step 3: 設定と照合してマーカーを修復
            if (_config.VirtualDesktops.Ids.TryGetValue("VD0", out string? vd0String) &&
                Guid.TryParse(vd0String, out Guid vd0Guid) &&
                currentGuid.Value == vd0Guid)
            {
                Debug.WriteLine("[FastDetector] 現在のデスクトップはVD0、マーカーを修復中...");
                RecreateMarker("VD0", vd0Guid);
                return "VD0";
            }
            else if (_config.VirtualDesktops.Ids.TryGetValue("VD1", out string? vd1String) &&
                     Guid.TryParse(vd1String, out Guid vd1Guid) &&
                     currentGuid.Value == vd1Guid)
            {
                Debug.WriteLine("[FastDetector] 現在のデスクトップはVD1、マーカーを修復中...");
                RecreateMarker("VD1", vd1Guid);
                return "VD1";
            }

            Debug.WriteLine($"[FastDetector] 未知のデスクトップGUID: {currentGuid}");
            return null;
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[FastDetector] リカバリー中にエラー: {ex.Message}");
            return null;
        }
    }

    public bool ValidateAndRepairAll()
    {
        if (_disposed || _markers.Count == 0)
            return false;

        try
        {
            bool allValid = true;
            var markersToRecreate = new List<(string desktopId, Guid guid)>();

            foreach (var kvp in _markers)
            {
                if (!kvp.Value.ValidatePosition())
                {
                    allValid = false;
                    
                    // 期待するGUIDで再作成をスケジュール
                    string? guidString = _config.VirtualDesktops.Ids.GetValueOrDefault(kvp.Key);
                    if (!string.IsNullOrEmpty(guidString) && Guid.TryParse(guidString, out Guid guid))
                    {
                        markersToRecreate.Add((kvp.Key, guid));
                    }
                }
            }

            // 無効なマーカーを再作成
            foreach (var (desktopId, guid) in markersToRecreate)
            {
                RecreateMarker(desktopId, guid);
            }

            return allValid;
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[FastDetector] マーカー検証中にエラー: {ex.Message}");
            return false;
        }
    }

    private void RecreateMarker(string desktopId, Guid desktopGuid)
    {
        try
        {
            // 古いマーカーを破棄
            if (_markers.TryGetValue(desktopId, out var oldMarker))
            {
                DisposeMarkerOnUiThread(oldMarker);
            }

            // 新しいマーカーを作成
            var newMarker = CreateMarkerOnUiThread(desktopId, desktopGuid);
            _markers[desktopId] = newMarker;

            Debug.WriteLine($"[FastDetector] {desktopId} のマーカーを再作成しました");
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[FastDetector] {desktopId} のマーカー再作成中にエラー: {ex.Message}");
        }
    }

    private void DisposeMarkers()
    {
        foreach (var marker in _markers.Values)
        {
            try
            {
                DisposeMarkerOnUiThread(marker);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[FastDetector] マーカー破棄中にエラー: {ex.Message}");
            }
        }
        _markers.Clear();
    }

    private DesktopMarkerWindow CreateMarkerOnUiThread(string desktopId, Guid desktopGuid)
    {
        if (_uiInvoker.InvokeRequired)
        {
            DesktopMarkerWindow? created = null;
            _uiInvoker.Invoke(new Action(() =>
            {
                created = new DesktopMarkerWindow(desktopId, desktopGuid, _vdManager);
            }));
            return created!;
        }
        else
        {
            return new DesktopMarkerWindow(desktopId, desktopGuid, _vdManager);
        }
    }

    private void DisposeMarkerOnUiThread(DesktopMarkerWindow marker)
    {
        if (_uiInvoker.InvokeRequired)
        {
            _uiInvoker.Invoke(new Action(() => marker.Dispose()));
        }
        else
        {
            marker.Dispose();
        }
    }

    public void Dispose()
    {
        if (!_disposed)
        {
            Debug.WriteLine("[FastDetector] FastDesktopDetectorを破棄中...");
            DisposeMarkers();
            _disposed = true;
        }
    }
}
