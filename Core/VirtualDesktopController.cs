using System;
using System.Collections.Generic;
using System.Drawing;
using TwoMiceVD.Configuration;
using TwoMiceVD.Interop;
// H.InputSimulator 1.5.0 exposes WindowsInput-compatible namespaces
using WindowsInput;
using static TwoMiceVD.Interop.NativeMethods;

namespace TwoMiceVD.Core;

public class VirtualDesktopController
{
    private int _currentDesktopIndex = 0;
    private readonly InputSimulator _inputSimulator;
    private readonly Dictionary<int, Point> _desktopCursorPositions = new Dictionary<int, Point>();

    public VirtualDesktopController(ConfigStore? config)
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
}
