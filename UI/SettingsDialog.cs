using System;
using System.Drawing;
using System.Windows.Forms;
using TwoMiceVD.Configuration;

namespace TwoMiceVD.UI;

public partial class SettingsDialog : Form
{
    private readonly ConfigStore _config;
    private TrackBar? _thresholdTrackBar;
    private TrackBar? _hysteresisTrackBar;
    private Label? _thresholdLabel;
    private Label? _hysteresisLabel;
    private CheckBox? _enableMousePositionCheckBox;
    private CheckBox? _exclusiveActiveCheckBox;
    private TrackBar? _activeHoldTrackBar;
    private Label? _activeHoldLabel;
    private CheckBox? _enableDeviceConnectionCheckBox;
    private Button? _okButton;
    private Button? _cancelButton;

    public SettingsDialog(ConfigStore config)
    {
        _config = config;
        InitializeComponent();
        LoadSettings();
    }

    private void InitializeComponent()
    {
        Text = "設定";
        Size = new Size(460, 440);
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;

        var thresholdTitleLabel = new Label
        {
            Text = "感度（移動しきい値）",
            Location = new Point(20, 20),
            Size = new Size(150, 20)
        };
        Controls.Add(thresholdTitleLabel);

        _thresholdLabel = new Label
        {
            Location = new Point(300, 20),
            Size = new Size(50, 20),
            TextAlign = ContentAlignment.MiddleRight
        };
        Controls.Add(_thresholdLabel);

        _thresholdTrackBar = new TrackBar
        {
            Location = new Point(20, 45),
            Size = new Size(330, 45),
            Minimum = 1,
            Maximum = 20,
            TickFrequency = 1
        };
        _thresholdTrackBar.ValueChanged += (s, e) => 
        {
            if (_thresholdLabel != null && _thresholdTrackBar != null)
                _thresholdLabel.Text = _thresholdTrackBar.Value.ToString();
        };
        Controls.Add(_thresholdTrackBar);

        var hysteresisTitleLabel = new Label
        {
            Text = "クールダウン時間（ミリ秒）",
            Location = new Point(20, 100),
            Size = new Size(150, 20)
        };
        Controls.Add(hysteresisTitleLabel);

        _hysteresisLabel = new Label
        {
            Location = new Point(300, 100),
            Size = new Size(50, 20),
            TextAlign = ContentAlignment.MiddleRight
        };
        Controls.Add(_hysteresisLabel);

        _hysteresisTrackBar = new TrackBar
        {
            Location = new Point(20, 125),
            Size = new Size(330, 45),
            Minimum = 100,
            Maximum = 1000,
            TickFrequency = 100
        };
        _hysteresisTrackBar.ValueChanged += (s, e) => 
        {
            if (_hysteresisLabel != null && _hysteresisTrackBar != null)
                _hysteresisLabel.Text = _hysteresisTrackBar.Value.ToString();
        };
        Controls.Add(_hysteresisTrackBar);

        _enableMousePositionCheckBox = new CheckBox
        {
            Text = "マウス位置記憶（マウス切り替え時に前の位置を記憶）",
            Location = new Point(20, 180),
            Size = new Size(350, 24),
            UseVisualStyleBackColor = true
        };
        Controls.Add(_enableMousePositionCheckBox);

        _exclusiveActiveCheckBox = new CheckBox
        {
            Text = "アクティブマウス優先モード（操作中は他方を無視）",
            Location = new Point(20, 210),
            Size = new Size(370, 24),
            UseVisualStyleBackColor = true
        };
        Controls.Add(_exclusiveActiveCheckBox);

        var activeHoldTitleLabel = new Label
        {
            Text = "アクティブ猶予（ミリ秒）",
            Location = new Point(20, 245),
            Size = new Size(180, 20)
        };
        Controls.Add(activeHoldTitleLabel);

        _activeHoldLabel = new Label
        {
            Location = new Point(300, 245),
            Size = new Size(50, 20),
            TextAlign = ContentAlignment.MiddleRight
        };
        Controls.Add(_activeHoldLabel);

        _activeHoldTrackBar = new TrackBar
        {
            Location = new Point(20, 270),
            Size = new Size(330, 45),
            Minimum = 50,
            Maximum = 1000,
            TickFrequency = 50
        };
        _activeHoldTrackBar.ValueChanged += (s, e) =>
        {
            if (_activeHoldLabel != null && _activeHoldTrackBar != null)
                _activeHoldLabel.Text = _activeHoldTrackBar.Value.ToString();
        };
        Controls.Add(_activeHoldTrackBar);

        _enableDeviceConnectionCheckBox = new CheckBox
        {
            Text = "接続状態監視（2台→1台で自動一時停止）［再起動後に反映］",
            Location = new Point(20, 320),
            Size = new Size(420, 24),
            UseVisualStyleBackColor = true
        };
        Controls.Add(_enableDeviceConnectionCheckBox);

        _okButton = new Button
        {
            Text = "OK",
            Location = new Point(260, 360),
            Size = new Size(75, 25),
            DialogResult = DialogResult.OK
        };
        _okButton.Click += (s, e) => SaveSettings();
        Controls.Add(_okButton);

        _cancelButton = new Button
        {
            Text = "キャンセル",
            Location = new Point(345, 360),
            Size = new Size(75, 25),
            DialogResult = DialogResult.Cancel
        };
        Controls.Add(_cancelButton);

        AcceptButton = _okButton;
        CancelButton = _cancelButton;
    }

    private void LoadSettings()
    {
        if (_thresholdTrackBar != null)
        {
            _thresholdTrackBar.Value = _config.ThresholdMovement;
        }
        
        if (_hysteresisTrackBar != null)
        {
            _hysteresisTrackBar.Value = _config.HysteresisMs;
        }
        
        if (_thresholdLabel != null)
        {
            _thresholdLabel.Text = _config.ThresholdMovement.ToString();
        }
        
        if (_hysteresisLabel != null)
        {
            _hysteresisLabel.Text = _config.HysteresisMs.ToString();
        }
        
        if (_enableMousePositionCheckBox != null)
        {
            _enableMousePositionCheckBox.Checked = _config.EnableMousePositionMemory;
        }

        if (_exclusiveActiveCheckBox != null)
        {
            _exclusiveActiveCheckBox.Checked = _config.ExclusiveActiveMouse;
        }

        if (_activeHoldTrackBar != null)
        {
            _activeHoldTrackBar.Value = Math.Max(50, Math.Min(1000, _config.ActiveHoldMs));
        }

        if (_activeHoldLabel != null)
        {
            _activeHoldLabel.Text = _config.ActiveHoldMs.ToString();
        }

        if (_enableDeviceConnectionCheckBox != null)
        {
            _enableDeviceConnectionCheckBox.Checked = _config.EnableDeviceConnectionCheck;
        }
    }

    private void SaveSettings()
    {
        if (_thresholdTrackBar != null)
        {
            _config.ThresholdMovement = _thresholdTrackBar.Value;
        }
        
        if (_hysteresisTrackBar != null)
        {
            _config.HysteresisMs = _hysteresisTrackBar.Value;
        }
        
        if (_enableMousePositionCheckBox != null)
        {
            _config.EnableMousePositionMemory = _enableMousePositionCheckBox.Checked;
        }

        if (_exclusiveActiveCheckBox != null)
        {
            _config.ExclusiveActiveMouse = _exclusiveActiveCheckBox.Checked;
        }

        if (_activeHoldTrackBar != null)
        {
            _config.ActiveHoldMs = _activeHoldTrackBar.Value;
        }

        if (_enableDeviceConnectionCheckBox != null)
        {
            _config.EnableDeviceConnectionCheck = _enableDeviceConnectionCheckBox.Checked;
        }

        _config.Save();
    }
}
