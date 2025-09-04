using System;
using System.Collections.Generic;
using System.Drawing;
using TwoMiceVD.Interop;
using static TwoMiceVD.Interop.NativeMethods;

namespace TwoMiceVD.Core;

/// <summary>
/// マウスごとのカーソル位置を管理するクラス
/// マウス切り替え時に前のマウスの位置を保存し、戻った時に復元する
/// </summary>
public class MousePositionManager
{
    private readonly Dictionary<string, Point> _devicePositions = new Dictionary<string, Point>();
    private string? _lastActiveDevice = null;
    private DateTime _lastRestoreTime = DateTime.MinValue;
    private const int RESTORE_COOLDOWN_MS = 200;
    private const int DISTANCE_THRESHOLD = 100;

    /// <summary>
    /// デバイスの移動を処理し、位置を更新して必要に応じて復元する
    /// </summary>
    /// <param name="deviceId">移動したデバイスID</param>
    /// <param name="deltaX">X軸移動量（使用しない）</param>
    /// <param name="deltaY">Y軸移動量（使用しない）</param>
    public void UpdateDevicePosition(string deviceId, int deltaX, int deltaY)
    {
        try
        {
            // 現在のカーソル位置を取得
            if (!GetCursorPos(out POINT currentPoint)) return;


            // デバイス切り替えを検知
            bool deviceChanged = _lastActiveDevice != null && _lastActiveDevice != deviceId;

            bool restored = false;
            Point restoredPos = default;

            if (deviceChanged)
            {
                // クールダウン期間チェック
                TimeSpan timeSinceLastRestore = DateTime.Now - _lastRestoreTime;
                if (timeSinceLastRestore.TotalMilliseconds > RESTORE_COOLDOWN_MS)
                {
                    // 復元候補を取得（存在する場合のみ復元を行う）
                    if (_devicePositions.TryGetValue(deviceId, out Point saved))
                    {
                        // 現在位置と保存位置の距離を確認し、十分離れていれば復元
                        int dist = Math.Abs(currentPoint.X - saved.X) + Math.Abs(currentPoint.Y - saved.Y);
                        if (dist > DISTANCE_THRESHOLD)
                        {
                            SetCursorPos(saved.X, saved.Y);
                            restored = true;
                            restoredPos = saved;
                        }
                    }

                    _lastRestoreTime = DateTime.Now;
                }
            }

            // 保存位置の更新
            // - 復元した場合は"復元した位置"をそのまま保持
            // - 復元していない場合は、現在のカーソル位置で更新
            if (restored)
            {
                _devicePositions[deviceId] = restoredPos;
            }
            else
            {
                _devicePositions[deviceId] = new Point(currentPoint.X, currentPoint.Y);
            }

            _lastActiveDevice = deviceId;
        }
        catch
        {
            // エラーは握り潰して安定性優先
        }
    }

    // 本クラスでは抑制や打ち消しは行わず、
    // Program 側でイベントの採否を判定する（単一カーソル前提）

    /// <summary>
    /// 指定されたデバイスの保存位置にカーソルを復元する
    /// </summary>
    /// <param name="deviceId">復元対象のデバイスID</param>
    /// <param name="currentPoint">現在のカーソル位置</param>
    private void RestorePosition(string deviceId, POINT currentPoint)
    {
        try
        {
            if (!_devicePositions.TryGetValue(deviceId, out Point savedPosition)) return;
            
            // 現在位置と保存位置の距離を計算
            int distance = Math.Abs(currentPoint.X - savedPosition.X) + 
                          Math.Abs(currentPoint.Y - savedPosition.Y);
            
            // 距離が閾値を超える場合のみ復元
            if (distance > DISTANCE_THRESHOLD)
            {
                SetCursorPos(savedPosition.X, savedPosition.Y);
                // デバッグ用（必要に応じてコメントアウト）
                System.Diagnostics.Debug.WriteLine(
                    $"位置復元: {deviceId} ({currentPoint.X},{currentPoint.Y}) -> ({savedPosition.X},{savedPosition.Y})");
            }
        }
        catch
        {
            // エラーは握り潰して安定性優先
        }
    }

    /// <summary>
    /// 指定されたデバイスの保存位置を取得する（デバッグ用）
    /// </summary>
    /// <param name="deviceId">デバイスID</param>
    /// <returns>保存されている位置、または null</returns>
    public Point? GetSavedPosition(string deviceId)
    {
        return _devicePositions.TryGetValue(deviceId, out Point position) ? position : null;
    }

    /// <summary>
    /// すべての保存位置をクリアする
    /// </summary>
    public void ClearAllPositions()
    {
        _devicePositions.Clear();
        _lastActiveDevice = null;
        _lastRestoreTime = DateTime.MinValue;
    }
    
    /// <summary>
    /// 現在の状態を取得する（デバッグ用）
    /// </summary>
    /// <returns>現在アクティブなデバイスと保存位置の数</returns>
    public (string? ActiveDevice, int SavedPositionCount) GetCurrentState()
    {
        return (_lastActiveDevice, _devicePositions.Count);
    }
}
