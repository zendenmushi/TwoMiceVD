using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using Newtonsoft.Json;

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
                swap_bindings = SwapBindings
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
        }
        
        if (data?.impl != null)
        {
            config.PreferInternalApi = data.impl.prefer_internal_api;
        }
        
        return config;
    }
}