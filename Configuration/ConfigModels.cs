using System.Collections.Generic;
using Newtonsoft.Json;

namespace TwoMiceVD.Configuration;

public class DeviceInfo
{
    [JsonProperty("device_path")]
    public string DevicePath { get; set; } = string.Empty;
    
    [JsonProperty("vid")]
    public int Vid { get; set; }
    
    [JsonProperty("pid")]
    public int Pid { get; set; }
}

public class VirtualDesktopInfo
{
    [JsonProperty("mode")]
    public string Mode { get; set; } = "GUID";
    
    [JsonProperty("ids")]
    public Dictionary<string, string> Ids { get; set; } = new Dictionary<string, string>();
    
    [JsonProperty("target_count")]
    public int TargetCount { get; set; } = 2;
}

internal class ConfigData
{
    public Dictionary<string, DeviceInfo>? devices { get; set; }
    public Dictionary<string, string>? bindings { get; set; }
    public VirtualDesktopInfo? virtual_desktops { get; set; }
    public SwitchConfig? @switch { get; set; }
    public UiConfig? ui { get; set; }
    public ImplementationConfig? impl { get; set; }
}

internal class SwitchConfig
{
    public int threshold { get; set; }
    public int hysteresis_ms { get; set; }
    public int max_per_sec { get; set; }
}

internal class UiConfig
{
    public bool startup_run { get; set; }
    public bool swap_bindings { get; set; }
}

internal class ImplementationConfig
{
    public bool prefer_internal_api { get; set; }
}