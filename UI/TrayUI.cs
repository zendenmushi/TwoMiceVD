using System;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;
using TwoMiceVD.Configuration;
using TwoMiceVD.Core;
using TwoMiceVD.Input;

namespace TwoMiceVD.UI;

public class TrayUI : IDisposable
{
    private readonly NotifyIcon _icon;
    private readonly ConfigStore _config;
    private readonly RawInputManager _rawInputManager;
    private ContextMenuStrip? _menu;
    private PairingDialog? _pairingDialog;
    private SettingsDialog? _settingsDialog;
    private readonly SwitchPolicy _switchPolicy;
    private readonly VirtualDesktopController _vdController;
    
    public event EventHandler? PairingRequested;

    public TrayUI(ConfigStore config, RawInputManager rawInputManager, SwitchPolicy switchPolicy, VirtualDesktopController vdController)
    {
        _config = config;
        _rawInputManager = rawInputManager;
        _switchPolicy = switchPolicy;
        _vdController = vdController;
        _menu = BuildContextMenu(); // 強参照を保持してGCによる解放を防ぐ
        _icon = new NotifyIcon
        {
            Text = "TwoMiceVD - 2マウス仮想デスクトップ切替",
            Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath),
            Visible = true,
            ContextMenuStrip = _menu
        };
        
        _icon.DoubleClick += (s, e) => ShowSettings();

        // Windows 10/11 で右クリックメニューが稀に出ない対策として手動表示も実装
        _icon.MouseUp += (s, e) =>
        {
            try
            {
                if (e.Button == MouseButtons.Right)
                {
                    // タスクトレイ位置にメニューを表示
                    _menu?.Show(Control.MousePosition);
                }
                else if (e.Button == MouseButtons.Left)
                {
                    // 左クリックでもメニューを出したい場合は以下を有効化
                    // _menu?.Show(Control.MousePosition);
                }
            }
            catch { /* 例外は握り潰して安定性優先 */ }
        };
    }

    private ContextMenuStrip BuildContextMenu()
    {
        var menu = new ContextMenuStrip();

        var pairItem = new ToolStripMenuItem("ペアリング開始")
        {
            ToolTipText = "マウスA/Bの識別を再設定します"
        };
        pairItem.Click += (s, e) => StartPairing();
        menu.Items.Add(pairItem);

        menu.Items.Add(new ToolStripSeparator());

        var swapItem = new ToolStripMenuItem("マウス割当反転")
        {
            Checked = _config.SwapBindings,
            CheckOnClick = true,
            ToolTipText = "マウスA/Bの割り当てを入れ替えます"
        };
        swapItem.Click += (s, e) => 
        {
            _config.SwapBindings = swapItem.Checked;
            _config.Save();
        };
        menu.Items.Add(swapItem);

        var settingsItem = new ToolStripMenuItem("設定")
        {
            ToolTipText = "感度、クールダウンなどを設定します"
        };
        settingsItem.Click += (s, e) => ShowSettings();
        menu.Items.Add(settingsItem);

        var startupItem = new ToolStripMenuItem("起動時に自動開始")
        {
            Checked = _config.StartupRun,
            CheckOnClick = true,
            ToolTipText = "Windows起動時にアプリを自動起動します"
        };
        startupItem.Click += (s, e) =>
        {
            _config.StartupRun = startupItem.Checked;
            SetStartupRegistration(startupItem.Checked);
            _config.Save();
        };
        menu.Items.Add(startupItem);

        menu.Items.Add(new ToolStripSeparator());

        var aboutItem = new ToolStripMenuItem("TwoMiceVD について");
        aboutItem.Click += (s, e) => ShowAbout();
        menu.Items.Add(aboutItem);

        var exitItem = new ToolStripMenuItem("終了");
        exitItem.Click += (s, e) => Application.Exit();
        menu.Items.Add(exitItem);

        return menu;
    }

    private async void StartPairing()
    {
        // ペアリング開始時に切り替えを無効化
        _switchPolicy.IsPairing = true;
        
        try
        {
            // 仮想デスクトップ構成の検証を実行
            var validationResult = await _vdController.ValidateDesktopConfigurationAsync();
            
            if (!validationResult.IsValid)
            {
                // エラータイプによって処理を分岐
                if (validationResult.ErrorType == VirtualDesktopController.DesktopValidationErrorType.NotEnoughDesktops)
                {
                    // デスクトップが1枚の場合：作成を提案
                    var result = MessageBox.Show(
                        "仮想デスクトップが1枚しかありません。\n\n" +
                        "新しいデスクトップを自動作成してペアリングを続行しますか？\n" +
                        "（Win+Ctrl+D ホットキーを送信して作成します）",
                        "仮想デスクトップの作成",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question
                    );

                    if (result == DialogResult.Yes)
                    {
                        try
                        {
                            // 新しいデスクトップを作成してGUIDも取得
                            var guids = await _vdController.CreateDesktopAsync();
                            if (guids == null)
                            {
                                MessageBox.Show(
                                    "デスクトップ作成またはGUID取得に失敗しました。",
                                    "エラー",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Error
                                );
                                return;
                            }
                            
                            // 検証結果を手動で構築（検証成功として扱う）
                            validationResult = new VirtualDesktopController.DesktopValidationResult
                            {
                                IsValid = true,
                                ErrorType = VirtualDesktopController.DesktopValidationErrorType.None,
                                VD0Guid = guids.Value.vd0,
                                VD1Guid = guids.Value.vd1
                            };
                        }
                        catch (Exception createEx)
                        {
                            MessageBox.Show(
                                $"デスクトップ作成中にエラーが発生しました:\n{createEx.Message}",
                                "作成エラー",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error
                            );
                            return;
                        }
                    }
                    else
                    {
                        // ユーザーがキャンセルした場合はペアリングを中止
                        return;
                    }
                }
                else
                {
                    // その他のエラーの場合：従来通りエラーメッセージを表示
                    MessageBox.Show(
                        validationResult.ErrorMessage ?? "仮想デスクトップ構成に問題があります。",
                        "ペアリング開始エラー",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    );
                    return;
                }
            }
            
            // 検証成功時：GUIDを設定に保存
            _config.VirtualDesktops.Mode = "GUID";
            _config.VirtualDesktops.Ids["VD0"] = validationResult.VD0Guid.ToString();
            _config.VirtualDesktops.Ids["VD1"] = validationResult.VD1Guid.ToString();
            _config.Save();
            
            // ペアリングダイアログを表示
            if (_pairingDialog == null || _pairingDialog.IsDisposed)
            {
                _pairingDialog = new PairingDialog(_config, _rawInputManager);
            }
            
            _pairingDialog.ShowDialog();
            
            PairingRequested?.Invoke(this, EventArgs.Empty);
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"ペアリング処理中にエラーが発生しました：{ex.Message}",
                "エラー",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            );
        }
        finally
        {
            // ペアリング終了時に切り替えを有効化
            _switchPolicy.IsPairing = false;
        }
    }

    private void ShowSettings()
    {
        if (_settingsDialog == null || _settingsDialog.IsDisposed)
        {
            _settingsDialog = new SettingsDialog(_config);
        }
        
        _settingsDialog.ShowDialog();
    }

    private void ShowAbout()
    {
        MessageBox.Show(
            "TwoMiceVD v1.0\n\n" +
            "2台の物理マウスを個別に認識し、各マウスに紐づく\n" +
            "仮想デスクトップへ自動で切り替えるアプリケーション\n\n" +
            "Windows 11 22H2 以降対応",
            "TwoMiceVD について",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information
        );
    }

    private void SetStartupRegistration(bool enable)
    {
        const string keyName = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run";
        const string valueName = "TwoMiceVD";

        try
        {
            using (var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(keyName, true))
            {
                if (enable)
                {
                    string? exePath = System.Diagnostics.Process.GetCurrentProcess().MainModule?.FileName;
                    if (!string.IsNullOrEmpty(exePath))
                    {
                        key?.SetValue(valueName, exePath);
                    }
                }
                else
                {
                    key?.DeleteValue(valueName, false);
                }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"起動設定の変更に失敗しました: {ex.Message}", "エラー", 
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    public void ShowNotification(string message, ToolTipIcon icon = ToolTipIcon.Info)
    {
        _icon.ShowBalloonTip(3000, "TwoMiceVD", message, icon);
    }

    public void Dispose()
    {
        if (_icon != null)
        {
            _icon.Visible = false;
            _icon.Dispose();
        }
        _pairingDialog?.Dispose();
        _settingsDialog?.Dispose();
        _menu?.Dispose();
    }
}