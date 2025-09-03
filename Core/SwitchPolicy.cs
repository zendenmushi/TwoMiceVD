using System;
using System.Collections.Generic;
using TwoMiceVD.Configuration;

namespace TwoMiceVD.Core;

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
            string? target = _config.GetDesktopIdForDevice(deviceId);
            if (!string.IsNullOrEmpty(target))
            {
                _controller.SwitchTo(target);
                _lastSwitchTime = DateTime.Now;
                _movementBuckets.Clear();
            }
        }
    }
}