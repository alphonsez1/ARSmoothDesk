using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace ARContentStabilizer
{
    public class ConfigManager
    {
        // Configuration file path
        private readonly string configFilePath;
        
        #region Configuration Properties

        // Display and content settings
        public Size DisplaySize { get; set; } = new Size(1920, 1080);
        public Size ContentSize { get; set; } = new Size(1366, 768);
        public int TargetDisplayIndex { get; set; } = -1; // Default to last display
        public int SourceDisplayIndex { get; set; } = 0;  // Default to primary display
        public float FieldOfView { get; set; } = 52.0f;   // Field of view in degrees

        // Movement settings
        public float FollowSpeedUp { get; set; } = 1.5f;
        public float FollowSpeedDown { get; set; } = 2.5f;
        public int CaptureFrameRate { get; set; } = 60;

        // Performance settings
        public bool UseDirectCapture { get; set; } = true;
        public bool UseLowQualityScaling { get; set; } = true;
        public bool UseThreadedCapture { get; set; } = true;
        public int SkipFrames { get; set; } = 0;
        public bool UseDoubleBuffering { get; set; } = true;
        public int CaptureThreadSleepTime { get; set; } = 5;

        #endregion

        public ConfigManager(string configPath = null)
        {
            configFilePath = configPath ?? "ARConfig.txt";
            LoadConfiguration();
        }

        public void LoadConfiguration()
        {
            try
            {
                if (File.Exists(configFilePath))
                {
                    string[] lines = File.ReadAllLines(configFilePath);
                    foreach (string line in lines)
                    {
                        string trimmedLine = line.Trim();
                        if (string.IsNullOrEmpty(trimmedLine) || trimmedLine.StartsWith("#"))
                            continue; // Skip comments and empty lines

                        string[] parts = trimmedLine.Split('=');
                        if (parts.Length != 2)
                            continue;

                        string key = parts[0].Trim();
                        string value = parts[1].Trim();

                        ParseConfigValue(key, value);
                    }
                }
                else
                {
                    // Create default config file if it doesn't exist
                    CreateDefaultConfigFile();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading configuration: {ex.Message}\nUsing default settings.",
                    "Configuration Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void ParseConfigValue(string key, string value)
        {
            switch (key.ToLower())
            {
                case "displaywidth":
                    if (int.TryParse(value, out int displayWidth))
                        DisplaySize = new Size(displayWidth, DisplaySize.Height);
                    break;
                case "displayheight":
                    if (int.TryParse(value, out int displayHeight))
                        DisplaySize = new Size(DisplaySize.Width, displayHeight);
                    break;
                case "contentwidth":
                    if (int.TryParse(value, out int contentWidth))
                        ContentSize = new Size(contentWidth, ContentSize.Height);
                    break;
                case "contentheight":
                    if (int.TryParse(value, out int contentHeight))
                        ContentSize = new Size(ContentSize.Width, contentHeight);
                    break;
                case "targetdisplay":
                    if (int.TryParse(value, out int targetDisplay))
                        TargetDisplayIndex = targetDisplay;
                    break;
                case "sourcedisplay":
                    if (int.TryParse(value, out int sourceDisplay))
                        SourceDisplayIndex = sourceDisplay;
                    break;
                case "followspeedup":
                    if (float.TryParse(value, out float speedUp))
                        FollowSpeedUp = speedUp;
                    break;
                case "followspeeddown":
                    if (float.TryParse(value, out float speedDown))
                        FollowSpeedDown = speedDown;
                    break;
                case "framerate":
                    if (int.TryParse(value, out int frameRate))
                        CaptureFrameRate = frameRate;
                    break;
                case "usedirectcapture":
                    if (bool.TryParse(value, out bool directCapture))
                        UseDirectCapture = directCapture;
                    break;
                case "uselowqualityscaling":
                    if (bool.TryParse(value, out bool lowQualityScaling))
                        UseLowQualityScaling = lowQualityScaling;
                    break;
                case "usethreadedcapture":
                    if (bool.TryParse(value, out bool threadedCapture))
                        UseThreadedCapture = threadedCapture;
                    break;
                case "skipframes":
                    if (int.TryParse(value, out int frames))
                        SkipFrames = frames;
                    break;
                case "fieldofview":
                    if (float.TryParse(value, out float fieldOfView))
                        FieldOfView = fieldOfView;
                    break;
                case "usedoublebuffering":
                    if (bool.TryParse(value, out bool doubleBuffering))
                        UseDoubleBuffering = doubleBuffering;
                    break;
                case "capturethreadsleeptime":
                    if (int.TryParse(value, out int sleepTime))
                        CaptureThreadSleepTime = sleepTime;
                    break;
            }
        }

        public void CreateDefaultConfigFile()
        {
            try
            {
                using (StreamWriter writer = new StreamWriter(configFilePath))
                {
                    writer.WriteLine("# AR Content Stabilizer Configuration");
                    writer.WriteLine("# Display sizes");
                    writer.WriteLine($"DisplayWidth={DisplaySize.Width}");
                    writer.WriteLine($"DisplayHeight={DisplaySize.Height}");
                    writer.WriteLine($"ContentWidth={ContentSize.Width}");
                    writer.WriteLine($"ContentHeight={ContentSize.Height}");
                    writer.WriteLine();
                    writer.WriteLine("# Display settings");
                    writer.WriteLine("# For target display: -1=last display, -2=second-to-last, 0=primary, 1,2,3...=specific displays");
                    writer.WriteLine($"TargetDisplay={TargetDisplayIndex}");
                    writer.WriteLine($"SourceDisplay={SourceDisplayIndex}");
                    writer.WriteLine();
                    writer.WriteLine("# Movement settings");
                    writer.WriteLine($"FollowSpeedUp={FollowSpeedUp}");
                    writer.WriteLine($"FollowSpeedDown={FollowSpeedDown}");
                    writer.WriteLine();
                    writer.WriteLine("# Capture settings");
                    writer.WriteLine($"FrameRate={CaptureFrameRate}");
                    writer.WriteLine($"FieldOfView={FieldOfView}");
                    writer.WriteLine();
                    writer.WriteLine("# Performance settings");
                    writer.WriteLine($"UseDirectCapture={UseDirectCapture}");
                    writer.WriteLine($"UseLowQualityScaling={UseLowQualityScaling}");
                    writer.WriteLine($"UseThreadedCapture={UseThreadedCapture}");
                    writer.WriteLine($"SkipFrames={SkipFrames}");
                    writer.WriteLine($"UseDoubleBuffering={UseDoubleBuffering}");
                    writer.WriteLine($"CaptureThreadSleepTime={CaptureThreadSleepTime}");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error creating default configuration: {ex.Message}",
                    "Configuration Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void SaveConfiguration()
        {
            CreateDefaultConfigFile(); // Reuse the same method to save current settings
        }
    }
}
