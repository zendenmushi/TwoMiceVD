using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;
using TwoMiceVD.Core;

namespace TwoMiceVD.Configuration;

public class ConfigStore
{
    private static readonly string ConfigDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "TwoMiceVD");
    private static readonly string ConfigFilePath = Path.Combine(ConfigDirectory, "config.json");

    public int ThresholdMovement { get; set; } = 5;
    public int HysteresisMs { get; set; } = 400;
    public int MaxSwitchPerSec { get; set; } = 2;
    public int VirtualDesktopTargetCount { get; set; } = 2;
    public bool PreferInternalApi { get; set; } = true;
    public bool StartupRun { get; set; } = false;
    public bool SwapBindings { get; set; } = false;
    public bool EnableMousePositionMemory { get; set; } = true;
    public bool ExclusiveActiveMouse { get; set; } = true;
    public int ActiveHoldMs { get; set; } = 150;

    public Dictionary<string, DeviceInfo> Devices { get; set; } = new Dictionary<string, DeviceInfo>();
    public Dictionary<string, string> Bindings { get; set; } = new Dictionary<string, string>();
    public VirtualDesktopInfo VirtualDesktops { get; set; } = new VirtualDesktopInfo();

    public static ConfigStore Load()
    {
        try
        {
            if (File.Exists(ConfigFilePath))
            {
                string json = File.ReadAllText(ConfigFilePath);
                var config = JsonConvert.DeserializeObject<ConfigData>(json);
                return FromConfigData(config);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"設定ファイルの読み込みに失敗しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        return new ConfigStore();
    }

    public void Save()
    {
        try
        {
            Directory.CreateDirectory(ConfigDirectory);
            var configData = ToConfigData();
            string json = JsonConvert.SerializeObject(configData, Formatting.Indented);
            File.WriteAllText(ConfigFilePath, json);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"設定ファイルの保存に失敗しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    public string? GetDesktopIdForDevice(string deviceId)
    {
        if (deviceId == null) return null;

        // デバイスID から デバイス識別子（A/B）を取得
        string? deviceKey = null;
        foreach (var device in Devices)
        {
            if (device.Value.DevicePath == deviceId)
            {
                deviceKey = device.Key;
                break;
            }
        }

        if (deviceKey == null) return null;

        // 割当の反転が有効な場合は反転させる
        if (SwapBindings)
        {
            deviceKey = deviceKey == "A" ? "B" : "A";
        }

        // デバイス識別子から仮想デスクトップIDを取得
        if (Bindings.TryGetValue(deviceKey, out string? desktopKey))
        {
            if (VirtualDesktops.Mode == "GUID" && VirtualDesktops.Ids.TryGetValue(desktopKey, out string? guid))
            {
                return guid;
            }
            return desktopKey; // インデックス形式の場合
        }

        return null;
    }

    public void SetDeviceBinding(string deviceKey, string devicePath, int vid, int pid)
    {
        Devices[deviceKey] = new DeviceInfo { DevicePath = devicePath, Vid = vid, Pid = pid };
    }

    public void SetDesktopBinding(string deviceKey, string desktopKey, string? desktopGuid = null)
    {
        Bindings[deviceKey] = desktopKey;
        if (!string.IsNullOrEmpty(desktopGuid))
        {
            VirtualDesktops.Mode = "GUID";
            VirtualDesktops.Ids[desktopKey] = desktopGuid;
        }
    }

    private ConfigData ToConfigData()
    {
        return new ConfigData
        {
            devices = Devices,
            bindings = Bindings,
            virtual_desktops = VirtualDesktops,
            @switch = new SwitchConfig
            {
                threshold = ThresholdMovement,
                hysteresis_ms = HysteresisMs,
                max_per_sec = MaxSwitchPerSec
            },
            ui = new UiConfig
            {
                startup_run = StartupRun,
                swap_bindings = SwapBindings,
                enable_mouse_position_memory = EnableMousePositionMemory,
                exclusive_active_mouse = ExclusiveActiveMouse,
                active_hold_ms = ActiveHoldMs
            },
            impl = new ImplementationConfig
            {
                prefer_internal_api = PreferInternalApi
            }
        };
    }

    private static ConfigStore FromConfigData(ConfigData? data)
    {
        var config = new ConfigStore();
        
        if (data?.devices != null)
            config.Devices = data.devices;
        
        if (data?.bindings != null)
            config.Bindings = data.bindings;
        
        if (data?.virtual_desktops != null)
            config.VirtualDesktops = data.virtual_desktops;
        
        if (data?.@switch != null)
        {
            config.ThresholdMovement = data.@switch.threshold;
            config.HysteresisMs = data.@switch.hysteresis_ms;
            config.MaxSwitchPerSec = data.@switch.max_per_sec;
        }
        
        if (data?.ui != null)
        {
            config.StartupRun = data.ui.startup_run;
            config.SwapBindings = data.ui.swap_bindings;
            config.EnableMousePositionMemory = data.ui.enable_mouse_position_memory;
            // 既存の設定ファイルにキーが無い場合は既定値を維持
            if (data.ui.exclusive_active_mouse.HasValue)
                config.ExclusiveActiveMouse = data.ui.exclusive_active_mouse.Value;
            if (data.ui.active_hold_ms.HasValue)
                config.ActiveHoldMs = data.ui.active_hold_ms.Value;
        }
        
        if (data?.impl != null)
        {
            config.PreferInternalApi = data.impl.prefer_internal_api;
        }
        
        return config;
    }

    /// <summary>
    /// 保存されたGUIDの有効性を検証
    /// </summary>
    public async Task<GuidValidationResult> ValidateStoredGuidsAsync(VirtualDesktopController vdController)
    {
        if (VirtualDesktops.Mode != "GUID")
        {
            return new GuidValidationResult
            {
                IsValid = false,
                Message = "GUID モードではありません。"
            };
        }

        if (!VirtualDesktops.Ids.TryGetValue("VD0", out string? vd0String) ||
            !VirtualDesktops.Ids.TryGetValue("VD1", out string? vd1String))
        {
            return new GuidValidationResult
            {
                IsValid = false,
                Message = "VD0 または VD1 の GUID が保存されていません。"
            };
        }

        if (!Guid.TryParse(vd0String, out Guid vd0Guid) ||
            !Guid.TryParse(vd1String, out Guid vd1Guid))
        {
            return new GuidValidationResult
            {
                IsValid = false,
                Message = "保存された GUID の形式が正しくありません。"
            };
        }

        // COM インターフェースを使用して GUID の存在を確認
        bool isValid = await vdController.ValidateStoredGuidsAsync(vd0Guid, vd1Guid);
        
        return new GuidValidationResult
        {
            IsValid = isValid,
            Message = isValid 
                ? "保存された仮想デスクトップ GUID は有効です。"
                : "保存された仮想デスクトップが見つかりません。Windows更新やデスクトップ削除により無効になった可能性があります。ペアリングをやり直してください。",
            VD0Guid = vd0Guid,
            VD1Guid = vd1Guid
        };
    }

    /// <summary>
    /// GUID を手動で設定する
    /// </summary>
    public void SaveDesktopGuids(Guid vd0Guid, Guid vd1Guid)
    {
        VirtualDesktops.Mode = "GUID";
        VirtualDesktops.Ids["VD0"] = vd0Guid.ToString();
        VirtualDesktops.Ids["VD1"] = vd1Guid.ToString();
        Save();
    }

    public class GuidValidationResult
    {
        public bool IsValid { get; set; }
        public string Message { get; set; } = string.Empty;
        public Guid VD0Guid { get; set; }
        public Guid VD1Guid { get; set; }
    }
}
