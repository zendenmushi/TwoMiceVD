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
        Size = new Size(400, 250);
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

        _okButton = new Button
        {
            Text = "OK",
            Location = new Point(220, 180),
            Size = new Size(75, 25),
            DialogResult = DialogResult.OK
        };
        _okButton.Click += (s, e) => SaveSettings();
        Controls.Add(_okButton);

        _cancelButton = new Button
        {
            Text = "キャンセル",
            Location = new Point(305, 180),
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
        
        _config.Save();
    }
}