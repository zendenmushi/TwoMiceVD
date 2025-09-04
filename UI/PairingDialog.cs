using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using TwoMiceVD.Configuration;
using TwoMiceVD.Core;
using TwoMiceVD.Input;

namespace TwoMiceVD.UI;

public partial class PairingDialog : Form
{
    private readonly ConfigStore _config;
    private readonly RawInputManager _sharedRawInput;
    private readonly Dictionary<string, string> _deviceNames = new Dictionary<string, string>();
    private bool _isCapturingA = false;
    private bool _isCapturingB = false;
    private Label? _instructionLabel;
    private Button? _captureAButton;
    private Button? _captureBButton;
    private Button? _completeButton;
    private Label? _statusLabel;

    public PairingDialog(ConfigStore config, RawInputManager sharedRawInput)
    {
        _config = config;
        _sharedRawInput = sharedRawInput;
        _sharedRawInput.DeviceMoved += OnDeviceMoved;
        InitializeComponent();
    }

    private void InitializeComponent()
    {
        Text = "マウスペアリング";
        Size = new Size(400, 300);
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;

        _instructionLabel = new Label
        {
            Text = "各マウスを順番に動かして識別してください",
            Location = new Point(20, 20),
            Size = new Size(360, 40),
            TextAlign = ContentAlignment.MiddleCenter
        };
        Controls.Add(_instructionLabel);

        _captureAButton = new Button
        {
            Text = "マウスA をキャプチャ",
            Location = new Point(50, 80),
            Size = new Size(120, 30)
        };
        _captureAButton.Click += (s, e) => StartCapture("A");
        Controls.Add(_captureAButton);

        _captureBButton = new Button
        {
            Text = "マウスB をキャプチャ",
            Location = new Point(230, 80),
            Size = new Size(120, 30)
        };
        _captureBButton.Click += (s, e) => StartCapture("B");
        Controls.Add(_captureBButton);

        _statusLabel = new Label
        {
            Text = "",
            Location = new Point(20, 130),
            Size = new Size(360, 60),
            TextAlign = ContentAlignment.MiddleCenter
        };
        Controls.Add(_statusLabel);

        _completeButton = new Button
        {
            Text = "完了",
            Location = new Point(160, 220),
            Size = new Size(80, 30),
            Enabled = false
        };
        _completeButton.Click += (s, e) => Complete();
        Controls.Add(_completeButton);
    }

    private void StartCapture(string deviceKey)
    {
        _isCapturingA = (deviceKey == "A");
        _isCapturingB = (deviceKey == "B");
        if (_statusLabel != null)
        {
            _statusLabel.Text = $"{deviceKey} を識別するため、そのマウスを動かしてください...";
        }
        
        if (_captureAButton != null) _captureAButton.Enabled = !_isCapturingA;
        if (_captureBButton != null) _captureBButton.Enabled = !_isCapturingB;
    }

    private void OnDeviceMoved(object? sender, DeviceMovedEventArgs e)
    {
        if (_isCapturingA)
        {
            _deviceNames["A"] = e.DeviceId;
            _isCapturingA = false;
            if (_statusLabel != null) _statusLabel.Text = "マウスA の識別が完了しました";
            if (_captureAButton != null) _captureAButton.Text = "マウスA ✓";
            if (_captureAButton != null) _captureAButton.Enabled = true;
            if (_captureBButton != null) _captureBButton.Enabled = true;
        }
        else if (_isCapturingB)
        {
            _deviceNames["B"] = e.DeviceId;
            _isCapturingB = false;
            if (_statusLabel != null) _statusLabel.Text = "マウスB の識別が完了しました";
            if (_captureBButton != null) _captureBButton.Text = "マウスB ✓";
            if (_captureAButton != null) _captureAButton.Enabled = true;
            if (_captureBButton != null) _captureBButton.Enabled = true;
        }

        if (_deviceNames.ContainsKey("A") && _deviceNames.ContainsKey("B"))
        {
            if (_completeButton != null) _completeButton.Enabled = true;
            if (_statusLabel != null) _statusLabel.Text = "両方のマウスが識別されました。完了ボタンを押してください。";
        }
    }

    private void Complete()
    {
        foreach (var device in _deviceNames)
        {
            _config.SetDeviceBinding(device.Key, device.Value, 0, 0);
            _config.SetDesktopBinding(device.Key, $"VD{(device.Key == "A" ? 0 : 1)}");
        }
        _config.Save();
        Close();
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _sharedRawInput.DeviceMoved -= OnDeviceMoved;
        }
        base.Dispose(disposing);
    }
}
