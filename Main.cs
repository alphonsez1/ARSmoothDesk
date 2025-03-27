using System;
using System.Drawing;
using System.Numerics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.IO;
using System.Drawing.Imaging;
using System.Drawing.Drawing2D;
using Timer = System.Windows.Forms.Timer;

namespace ARContentStabilizer
{
    public partial class MainForm : Form
    {
        #region DLL Imports

        [DllImport("shcore.dll")]
        static extern int SetProcessDpiAwareness(int awareness);

        // Constants for DPI awareness
        private const int PROCESS_PER_MONITOR_DPI_AWARE = 2;
        private const uint MONITOR_DEFAULTTONEAREST = 0x00000002;
        private const int MDT_EFFECTIVE_DPI = 0;

        #endregion

        #region Configuration and Settings

        // Configuration manager
        private readonly ConfigManager config;
        
        #endregion

        #region UI Components and Rendering

        // Content panel for rendering
        private PictureBox contentDisplay;
        
        // Screen capture manager
        private ScreenCaptureManager screenCaptureManager;
        
        #endregion

        #region Tracking and Movement

        // Head tracking manager
        private HeadTrackingManager headTracking;

        // Content position variables
        private PointF currentPosition;

        // Timer for update loop
        private Timer updateTimer;

        // Boundaries for content to stay visible
        private PointF minBoundary;
        private PointF maxBoundary;

        #endregion

        public MainForm()
        {
            // Initialize configuration manager
            config = new ConfigManager();
            
            // Make application DPI aware
            SetDpiAwareness();

            // Make the form borderless and fullscreen 
            this.FormBorderStyle = FormBorderStyle.None;
            this.WindowState = FormWindowState.Maximized;
            
            // Set background color to black
            this.BackColor = Color.Black;
            
            float radius = SphereTracking.CalculateDistanceToDisplay(
                config.DisplaySize.Width, config.DisplaySize.Height, config.FieldOfView);
            
            // Initialize head tracking
            headTracking = new HeadTrackingManager(radius);
            
            // Set the form to run on the target display
            SetTargetDisplay();
            
            // Calculate boundaries to keep content within display area
            UpdateBoundaries();
            
            // Initialize components
            InitializeComponents();
            
            // Wire up events
            this.FormClosing += MainForm_FormClosing;
            this.KeyDown += MainForm_KeyDown;
            this.Resize += MainForm_Resize;
            
            // Set up timer for update loop
            SetupTimer();
            
            // Initialize tracking system
            InitializeTracking();
        }

        #region Initialization Methods

        private void SetDpiAwareness()
        {
            try
            {
                // Make application per-monitor DPI aware
                SetProcessDpiAwareness(PROCESS_PER_MONITOR_DPI_AWARE);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to set DPI awareness: {ex.Message}");
            }
        }

        private void UpdateBoundaries()
        {
            // Calculate proper boundaries based on the form size, not the display size
            // This ensures the content can move within the current window
            minBoundary = new PointF(
                -(this.ClientSize.Width - config.ContentSize.Width) / 2f,
                -(this.ClientSize.Height - config.ContentSize.Height) / 2f
            );
            maxBoundary = new PointF(
                (this.ClientSize.Width - config.ContentSize.Width) / 2f,
                (this.ClientSize.Height - config.ContentSize.Height) / 2f
            );
            
            Console.WriteLine($"Boundary updated: Min({minBoundary.X}, {minBoundary.Y}), Max({maxBoundary.X}, {maxBoundary.Y})");
            
            // Update boundaries in head tracking manager
            if (headTracking != null)
            {
                headTracking.SetBoundaries(minBoundary, maxBoundary);
            }
        }

        private void SetTargetDisplay()
        {
            // Convert relative indices to absolute
            int displayCount = Screen.AllScreens.Length;
            int actualDisplayIndex = config.TargetDisplayIndex;
            
            if (config.TargetDisplayIndex < 0)
            {
                actualDisplayIndex = displayCount + config.TargetDisplayIndex;
                
                // Make sure we don't go below 0
                if (actualDisplayIndex < 0)
                    actualDisplayIndex = 0;
            }
            else if (config.TargetDisplayIndex >= displayCount)
            {
                actualDisplayIndex = displayCount - 1;
            }
            
            // Set the form's location to the target screen
            Screen targetScreen = Screen.AllScreens[actualDisplayIndex];
            this.StartPosition = FormStartPosition.Manual;
            this.Location = targetScreen.Bounds.Location;
            this.Size = targetScreen.Bounds.Size;
        }

        private void InitializeComponents()
        {
            // Create main picturebox to display captured content
            contentDisplay = new PictureBox
            {
                Size = config.ContentSize,
                SizeMode = PictureBoxSizeMode.StretchImage,
                BorderStyle = BorderStyle.None
            };
            
            // Enable double buffering on the PictureBox for smoother updates
            if (config.UseDoubleBuffering)
            {
                // Enable double buffering using reflection
                typeof(PictureBox).InvokeMember("DoubleBuffered", 
                    System.Reflection.BindingFlags.SetProperty | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic,
                    null, contentDisplay, new object[] { true });
            }
            
            // Add to form
            this.Controls.Add(contentDisplay);
            
            // Center initially
            currentPosition = new PointF(0, 0);
            UpdateContentPosition();

            // Initialize the screen capture manager
            screenCaptureManager = new ScreenCaptureManager(config, contentDisplay);
        }

        private void SetupTimer()
        {
            // Set up timer for movement update
            updateTimer = new Timer
            {
                Interval = 5 // ~100fps (movement should be smoother)
            };
            updateTimer.Tick += UpdateTimer_Tick;
        }

        private void InitializeTracking()
        {
            // Connect to AR glasses using the head tracking manager
            if (headTracking.Initialize())
            {
                // Start update loops
                updateTimer.Start();
                
                // Start screen capture
                screenCaptureManager.Start();
            }
            else
            {
                MessageBox.Show("Failed to connect to AR glasses. Using keyboard simulation.",
                    "Connection Failed", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                // Exit the application if connection fails
                this.Close();
                return;
            }
        }

        #endregion

        #region Event Handlers

        private void MainForm_KeyDown(object sender, KeyEventArgs e)
        {
            // Escape to exit
            if (e.KeyCode == Keys.Escape)
            {
                this.Close();
                return;
            }
        }

        private void UpdateTimer_Tick(object sender, EventArgs e)
        {
            var stopWatch = System.Diagnostics.Stopwatch.StartNew();

            Console.WriteLine("----------");

            // Update rotation data and calculate new position
            headTracking.UpdateRotationData();

            // Log tracking rays for debugging
            Console.WriteLine($"viewCenterRay: {headTracking.ViewFrontRay.X}, {headTracking.ViewFrontRay.Y}, {headTracking.ViewFrontRay.Z}");
            Console.WriteLine($"viewRightRay: {headTracking.ViewRightRay.X}, {headTracking.ViewRightRay.Y}, {headTracking.ViewRightRay.Z}");
            Console.WriteLine($"screenCenterRay before update: {headTracking.ScreenCenterRay.X}, {headTracking.ScreenCenterRay.Y}, {headTracking.ScreenCenterRay.Z}");

            // Get current position before update
            Console.WriteLine($"Current Position Before Update: {currentPosition.X}, {currentPosition.Y}");

            // Get new position from head tracking (includes clamping and ray adjustments)
            currentPosition = headTracking.UpdateHeadPosition();

            // Log the final position
            Console.WriteLine($"Current Position After Update: {currentPosition.X}, {currentPosition.Y}");

            // Update content position on screen
            UpdateContentPosition();

            // Print the time taken in milliseconds for debugging
            Console.WriteLine($"UpdateTimer_Tick time: {stopWatch.ElapsedMilliseconds} ms");
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            // Stop timers
            updateTimer.Stop();
            
            // Stop screen capture manager
            screenCaptureManager.Stop();
            
            // Disconnect from AR glasses
            headTracking.Shutdown();
        }

        private void MainForm_Resize(object sender, EventArgs e)
        {
            // Update boundaries when form is resized
            UpdateBoundaries();
            // Update content position based on new boundaries
            UpdateContentPosition();
        }

        #endregion

        #region Core Functionality

        private void UpdateContentPosition()
        {
            // Calculate the centered position
            float centerX = (this.ClientSize.Width - contentDisplay.Width) / 2.0f;
            float centerY = (this.ClientSize.Height - contentDisplay.Height) / 2.0f;
            
            // Apply the offset to our centered position
            int newX = (int)(centerX + currentPosition.X);
            int newY = (int)(centerY - currentPosition.Y);
            
            // Debug output to verify position calculation
            Console.WriteLine($"Content Position: ({newX}, {newY})");

            // Set the new location
            contentDisplay.Location = new Point(newX, newY);
        }

        #endregion

        [STAThread]
        static void Main()
        {
            // Set DLL directory to ensure dependencies can be found
            SetDllDirectory(Path.GetDirectoryName(Application.ExecutablePath));
            
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }

        // Add this to help find DLL dependencies
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool SetDllDirectory(string lpPathName);
    }
}