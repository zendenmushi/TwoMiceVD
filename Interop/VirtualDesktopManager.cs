using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace TwoMiceVD.Interop
{
    [ComImport]
    [Guid("A5CD92FF-29BE-454C-8D04-D82879FB3F1B")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IVirtualDesktopManager
    {
        int IsWindowOnCurrentVirtualDesktop(IntPtr topLevelWindow, out bool onCurrentDesktop);
        int GetWindowDesktopId(IntPtr topLevelWindow, out Guid desktopId);
        int MoveWindowToDesktop(IntPtr topLevelWindow, ref Guid desktopId);
    }

    [ComImport]
    [Guid("AA509086-5CA9-4C25-8F95-589D3C07B48A")]
    public class VirtualDesktopManager
    {
    }

    public class VirtualDesktopManagerWrapper : IDisposable
    {
        private IVirtualDesktopManager? _manager;
        private bool _disposed = false;

        public VirtualDesktopManagerWrapper()
        {
            try
            {
                _manager = (IVirtualDesktopManager)new VirtualDesktopManager();
            }
            catch (COMException ex)
            {
                Debug.WriteLine($"[TwoMiceVD] VirtualDesktopManager初期化失敗: HRESULT=0x{ex.HResult:X8}, Message={ex.Message}");
                throw new InvalidOperationException($"仮想デスクトップCOMインターフェースの初期化に失敗しました: {ex.Message}", ex);
            }
        }

        public bool IsAvailable => _manager != null && !_disposed;

        public Guid? GetWindowDesktopId(IntPtr hwnd)
        {
            if (_disposed || _manager == null)
                return null;

            try
            {
                int result = _manager.GetWindowDesktopId(hwnd, out Guid desktopId);
                return result == 0 ? desktopId : null; // S_OK = 0
            }
            catch (COMException ex)
            {
                if (ex.HResult == unchecked((int)0x8002802B)) // TYPE_E_ELEMENTNOTFOUND
                {
                    Debug.WriteLine($"[TwoMiceVD] GetWindowDesktopId: 指定されたデスクトップGUIDが見つかりません: HWND=0x{hwnd:X}, HRESULT=0x{ex.HResult:X8}");
                }
                else
                {
                    Debug.WriteLine($"[TwoMiceVD] GetWindowDesktopId失敗: HWND=0x{hwnd:X}, HRESULT=0x{ex.HResult:X8}, Message={ex.Message}");
                }
                return null;
            }
        }

        public bool MoveWindowToDesktop(IntPtr hwnd, Guid desktopId)
        {
            if (_disposed || _manager == null)
                return false;

            try
            {
                var guid = desktopId;
                int result = _manager.MoveWindowToDesktop(hwnd, ref guid);
                return result == 0; // S_OK = 0
            }
            catch (COMException ex)
            {
                if (ex.HResult == unchecked((int)0x8002802B)) // TYPE_E_ELEMENTNOTFOUND
                {
                    Debug.WriteLine($"[TwoMiceVD] MoveWindowToDesktop: 指定されたデスクトップGUID({desktopId})が見つかりません: HWND=0x{hwnd:X}, HRESULT=0x{ex.HResult:X8}");
                }
                else
                {
                    Debug.WriteLine($"[TwoMiceVD] MoveWindowToDesktop失敗: HWND=0x{hwnd:X}, DesktopGUID={desktopId}, HRESULT=0x{ex.HResult:X8}, Message={ex.Message}");
                }
                return false;
            }
        }

        public bool IsWindowOnCurrentVirtualDesktop(IntPtr hwnd)
        {
            if (_disposed || _manager == null)
                return false;

            try
            {
                int result = _manager.IsWindowOnCurrentVirtualDesktop(hwnd, out bool onCurrentDesktop);
                return result == 0 && onCurrentDesktop;
            }
            catch (COMException ex)
            {
                Debug.WriteLine($"[TwoMiceVD] IsWindowOnCurrentVirtualDesktop失敗: HWND=0x{hwnd:X}, HRESULT=0x{ex.HResult:X8}, Message={ex.Message}");
                return false;
            }
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                if (_manager != null)
                {
                    Marshal.ReleaseComObject(_manager);
                    _manager = null;
                }
                _disposed = true;
            }
        }
    }
}