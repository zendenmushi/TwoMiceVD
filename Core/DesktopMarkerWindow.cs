using System;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
using TwoMiceVD.Interop;

namespace TwoMiceVD.Core;

public class DesktopMarkerWindow : IDisposable
{
    private Form? _markerWindow;
    private Guid _expectedDesktopGuid;
    private readonly string _desktopId;
    private readonly VirtualDesktopManagerWrapper _vdManager;
    private bool _disposed = false;

    public DesktopMarkerWindow(string desktopId, Guid desktopGuid, VirtualDesktopManagerWrapper vdManager)
    {
        _desktopId = desktopId;
        _expectedDesktopGuid = desktopGuid;
        _vdManager = vdManager;

        CreateMarkerWindow();
        MoveToDesktop(desktopGuid);
    }

    public IntPtr Handle => _markerWindow?.Handle ?? IntPtr.Zero;
    public string DesktopId => _desktopId;
    public Guid ExpectedDesktopGuid => _expectedDesktopGuid;

    private void CreateMarkerWindow()
    {
        try
        {
            _markerWindow = new MarkerForm
            {
                Text = $"TwoMiceVD_Marker_{_desktopId}",
                FormBorderStyle = FormBorderStyle.None,
                ShowInTaskbar = false, // 後でExStyleを直接調整
                WindowState = FormWindowState.Normal,
                Size = new Size(1, 1),
                StartPosition = FormStartPosition.Manual,
                Location = new Point(-10000, -10000), // 画面外に配置
                Visible = false,
                TopMost = false
            };

            // ハンドルを確実に作成し、一度だけ表示してVDMに認識させる
            var handle = _markerWindow.Handle;
            Application.DoEvents();

            // 強制的に WS_EX_TOOLWINDOW を外し、スタイルを反映
            const int WS_EX_TOOLWINDOW = 0x00000080;
            int ex = NativeMethods.GetWindowLong(handle, NativeMethods.GWL_EXSTYLE);
            int exNew = ex & ~WS_EX_TOOLWINDOW;
            if (exNew != ex)
            {
                NativeMethods.SetWindowLong(handle, NativeMethods.GWL_EXSTYLE, exNew);
                NativeMethods.SetWindowPos(handle, IntPtr.Zero, 0, 0, 0, 0,
                    (uint)(NativeMethods.SWP_NOMOVE | NativeMethods.SWP_NOSIZE | NativeMethods.SWP_NOZORDER | NativeMethods.SWP_FRAMECHANGED));
            }

            // 常時可視（オフスクリーン）のまま維持してVDMの対象にする
            _markerWindow.Visible = true;
            Application.DoEvents();

            Debug.WriteLine($"[Marker] {_desktopId} のマーカーウィンドウを作成: Handle=0x{handle:X}");
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[Marker] {_desktopId} のマーカーウィンドウ作成に失敗: {ex.Message}");
            throw;
        }
    }

    private class MarkerForm : Form
    {
        // Win32 拡張スタイル定数
        private const int WS_EX_TOOLWINDOW = 0x00000080;
        private const int WS_EX_APPWINDOW = 0x00040000;
        private const int WS_EX_NOACTIVATE = 0x08000000;

        protected override CreateParams CreateParams
        {
            get
            {
                var cp = base.CreateParams;
                // ツールウィンドウ属性を外し、アプリウィンドウとして扱わせる
                cp.ExStyle &= ~WS_EX_TOOLWINDOW;
                cp.ExStyle |= WS_EX_APPWINDOW;
                // アクティブ化は不要
                cp.ExStyle |= WS_EX_NOACTIVATE;
                return cp;
            }
        }
    }

    private bool MoveToDesktop(Guid desktopGuid)
    {
        if (_markerWindow == null)
            return false;

        try
        {
            bool success = _vdManager.MoveWindowToDesktop(_markerWindow.Handle, desktopGuid);
            if (success)
            {
                Debug.WriteLine($"[Marker] {_desktopId} のマーカーをデスクトップ {desktopGuid} に配置");
            }
            else
            {
                Debug.WriteLine($"[Marker] {_desktopId} のマーカー配置に失敗: GUID={desktopGuid}");
            }
            return success;
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[Marker] {_desktopId} のマーカー配置中にエラー: {ex.Message}");
            return false;
        }
    }

    public bool ValidatePosition()
    {
        if (_disposed || _markerWindow == null)
            return false;

        try
        {
            // スレッドセーフなハンドル取得
            IntPtr handle = IntPtr.Zero;
            bool isDisposed = false;
            
            if (_markerWindow.InvokeRequired)
            {
                _markerWindow.Invoke(new Action(() => {
                    isDisposed = _markerWindow.IsDisposed;
                    if (!isDisposed)
                        handle = _markerWindow.Handle;
                }));
            }
            else
            {
                isDisposed = _markerWindow.IsDisposed;
                if (!isDisposed)
                    handle = _markerWindow.Handle;
            }

            if (isDisposed || handle == IntPtr.Zero)
            {
                Debug.WriteLine($"[Marker] {_desktopId} のマーカーウィンドウが破棄されています");
                return false;
            }

            // 現在配置されているデスクトップのGUIDを取得
            Guid? actualGuid = _vdManager.GetWindowDesktopId(handle);

            if (actualGuid == null)
            {
                Debug.WriteLine($"[Marker] {_desktopId} のマーカーウィンドウが見つかりません");
                return false;
            }

            if (actualGuid.Value != _expectedDesktopGuid)
            {
                Debug.WriteLine($"[Marker] {_desktopId} のマーカーが期待と異なるデスクトップにあります");
                Debug.WriteLine($"  期待: {_expectedDesktopGuid}");
                Debug.WriteLine($"  実際: {actualGuid.Value}");
                return false;
            }

            return true;
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[Marker] {_desktopId} の位置検証中にエラー: {ex.Message}");
            return false;
        }
    }

    public bool Reposition(Guid newDesktopGuid)
    {
        if (_disposed || _markerWindow == null || _markerWindow.IsDisposed)
            return false;

        try
        {
            _expectedDesktopGuid = newDesktopGuid;
            bool success = MoveToDesktop(newDesktopGuid);
            
            if (success)
            {
                Debug.WriteLine($"[Marker] {_desktopId} のマーカーを {newDesktopGuid} に再配置しました");
            }
            
            return success;
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[Marker] {_desktopId} のマーカー再配置中にエラー: {ex.Message}");
            return false;
        }
    }

    public bool IsOnCurrentDesktop()
    {
        if (_disposed || _markerWindow == null)
            return false;

        try
        {
            // スレッドセーフなハンドル取得
            IntPtr handle = IntPtr.Zero;
            bool isDisposed = false;
            
            if (_markerWindow.InvokeRequired)
            {
                _markerWindow.Invoke(new Action(() => {
                    isDisposed = _markerWindow.IsDisposed;
                    if (!isDisposed)
                        handle = _markerWindow.Handle;
                }));
            }
            else
            {
                isDisposed = _markerWindow.IsDisposed;
                if (!isDisposed)
                    handle = _markerWindow.Handle;
            }

            if (isDisposed || handle == IntPtr.Zero)
                return false;

            return _vdManager.IsWindowOnCurrentVirtualDesktop(handle);
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[Marker] {_desktopId} の現在デスクトップ判定中にエラー: {ex.Message}");
            return false;
        }
    }

    public void Dispose()
    {
        if (!_disposed)
        {
            try
            {
                if (_markerWindow != null && !_markerWindow.IsDisposed)
                {
                    _markerWindow.Dispose();
                    Debug.WriteLine($"[Marker] {_desktopId} のマーカーウィンドウを破棄しました");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[Marker] {_desktopId} のマーカー破棄中にエラー: {ex.Message}");
            }
            finally
            {
                _markerWindow = null;
                _disposed = true;
            }
        }
    }
}
