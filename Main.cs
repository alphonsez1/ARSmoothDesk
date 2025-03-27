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

        // AirAPI DLL Imports
        [DllImport("deps\\AirAPI_Windows.dll", CallingConvention = CallingConvention.Cdecl)]
        static extern int StartConnection();

        [DllImport("deps\\AirAPI_Windows.dll", CallingConvention = CallingConvention.Cdecl)]
        static extern int StopConnection();

        [DllImport("deps\\AirAPI_Windows.dll", CallingConvention = CallingConvention.Cdecl)]
        static extern IntPtr GetQuaternion();

        [DllImport("deps\\AirAPI_Windows.dll", CallingConvention = CallingConvention.Cdecl)]
        static extern IntPtr GetEuler();

        // Screen capture DLL imports
        [DllImport("user32.dll")]
        static extern IntPtr GetDC(IntPtr hwnd);

        [DllImport("user32.dll")]
        static extern int ReleaseDC(IntPtr hwnd, IntPtr hdc);

        [DllImport("gdi32.dll")]
        static extern bool BitBlt(IntPtr hdcDest, int nXDest, int nYDest, int nWidth, int nHeight, 
            IntPtr hdcSrc, int nXSrc, int nYSrc, int dwRop);

        [DllImport("gdi32.dll")]
        static extern IntPtr CreateCompatibleDC(IntPtr hdc);

        [DllImport("gdi32.dll")]
        static extern IntPtr CreateCompatibleBitmap(IntPtr hdc, int nWidth, int nHeight);

        [DllImport("gdi32.dll")]
        static extern IntPtr SelectObject(IntPtr hdc, IntPtr hgdiobj);

        [DllImport("gdi32.dll")]
        static extern bool DeleteObject(IntPtr hObject);

        [DllImport("gdi32.dll")]
        static extern bool DeleteDC(IntPtr hdc);

        [StructLayout(LayoutKind.Sequential)]
        struct POINT
        {
            public int x;
            public int y;
        }

        [StructLayout(LayoutKind.Sequential)]
        struct CURSORINFO
        {
            public int cbSize;
            public int flags;
            public IntPtr hCursor;
            public POINT ptScreenPos;
        }

        [StructLayout(LayoutKind.Sequential)]
        struct ICONINFO
        {
            public bool fIcon;
            public int xHotspot;
            public int yHotspot;
            public IntPtr hbmMask;
            public IntPtr hbmColor;
        }

        // Mouse cursor capture imports
        [DllImport("user32.dll")]
        static extern bool GetCursorInfo(out CURSORINFO pci);

        [DllImport("user32.dll")]
        static extern bool GetIconInfo(IntPtr hIcon, out ICONINFO piconinfo);

        [DllImport("user32.dll")]
        static extern IntPtr GetCursor();

        [DllImport("user32.dll")]
        static extern bool GetCursorPos(out POINT lpPoint);

        [DllImport("user32.dll")]
        static extern IntPtr CopyIcon(IntPtr hIcon);

        [DllImport("user32.dll")]
        static extern bool DrawIcon(IntPtr hdc, int x, int y, IntPtr hIcon);

        [DllImport("user32.dll")]
        static extern IntPtr MonitorFromWindow(IntPtr hwnd, uint dwFlags);

        [DllImport("user32.dll")]
        static extern bool GetDpiForMonitor(IntPtr hmonitor, int dpiType, out uint dpiX, out uint dpiY);

        [DllImport("shcore.dll")]
        static extern int SetProcessDpiAwareness(int awareness);

        // Constants for cursor
        private const int CURSOR_SHOWING = 0x00000001;

        // Constants for BitBlt
        private const int SRCCOPY = 0xCC0020;

        // Constants for DPI awareness
        private const int PROCESS_PER_MONITOR_DPI_AWARE = 2;
        private const uint MONITOR_DEFAULTTONEAREST = 0x00000002;
        private const int MDT_EFFECTIVE_DPI = 0;

        #endregion

        #region Configuration and Settings

        // Configuration manager
        private readonly ConfigManager config;
        
        // Scaling factors
        private float sourceScaleFactor = 1.0f;

        // Performance settings
        private Rectangle captureRegion = Rectangle.Empty; // Optional region to capture instead of full screen
        private bool captureInProgress = false;  // Flag to prevent overlapping captures
        private System.Threading.Thread captureThread; // Background thread for capture
        private volatile bool shutdownThreads = false; // Signal to shutdown background threads

        #endregion

        #region UI Components and Rendering

        // Content panel for rendering
        private PictureBox contentDisplay;
        
        // Bitmap for screen capture
        private Bitmap capturedScreen;
        private Bitmap scaledCapture;
        
        #endregion

        #region Tracking and Movement

        // Head tracking manager
        private HeadTrackingManager headTracking;

        // Content position variables
        private PointF currentPosition;

        // Timer for update loop
        private Timer updateTimer;
        private Timer captureTimer;

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
            this.BackColor = Color.White;
            
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
            
            // Set up timers for update loop and screen capture
            SetupTimers();
            
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
            
            // Create initial bitmaps
            capturedScreen = new Bitmap(Screen.AllScreens[GetSourceDisplayIndex()].Bounds.Width, 
                                       Screen.AllScreens[GetSourceDisplayIndex()].Bounds.Height);
            scaledCapture = new Bitmap(config.ContentSize.Width, config.ContentSize.Height);
            
            // Add to form
            this.Controls.Add(contentDisplay);
            
            // Center initially
            currentPosition = new PointF(0, 0);
            UpdateContentPosition();

            // Detect source display scaling factor
            UpdateSourceDisplayScaleFactor();
        }

        private void SetupTimers()
        {
            // Set up timer for movement update
            updateTimer = new Timer
            {
                Interval = 5 // ~100fps (movement should be smoother)
            };
            updateTimer.Tick += UpdateTimer_Tick;
            
            // Set threaded capture or timer-based capture
            if (config.UseThreadedCapture)
            {
                // Use a dedicated thread for capturing at maximum possible rate
                captureThread = new System.Threading.Thread(CaptureThreadMethod);
                captureThread.IsBackground = true;
                captureThread.Priority = System.Threading.ThreadPriority.AboveNormal;
                captureThread.Start();
            }
            else
            {
                // Set up timer for screen capture
                captureTimer = new Timer
                {
                    Interval = 1000 / config.CaptureFrameRate // Based on configured frame rate
                };
                captureTimer.Tick += CaptureTimer_Tick;
            }
        }
        
        private void CaptureThreadMethod()
        {
            // Setup for frame skipping
            int frameCounter = 0;
            
            try
            {
                while (!shutdownThreads)
                {
                    // Skip frames if configured
                    frameCounter++;
                    if (config.SkipFrames > 0 && frameCounter % (config.SkipFrames + 1) != 0)
                    {
                        System.Threading.Thread.Sleep(config.CaptureThreadSleepTime);
                        continue;
                    }
                    
                    // Check if previous capture is still processing
                    if (!captureInProgress)
                    {
                        captureInProgress = true;

                        try 
                        {
                            // Capture screen and update image
                            var stopWatch = System.Diagnostics.Stopwatch.StartNew();
                            CaptureScreen();
                            UpdateContentImage();
                            
                            // Add a small sleep to prevent using 100% CPU
                            // Calculate sleep time based on target frame rate
                            long elapsedMs = stopWatch.ElapsedMilliseconds;
                            int targetFrameTime = 1000 / config.CaptureFrameRate;
                            int sleepTime = Math.Max(config.CaptureThreadSleepTime, targetFrameTime - (int)elapsedMs);
                            
                            System.Threading.Thread.Sleep(sleepTime);
                        }
                        finally 
                        {
                            captureInProgress = false;
                        }
                    }
                    else
                    {
                        // Previous capture still processing, sleep for a short time
                        System.Threading.Thread.Sleep(config.CaptureThreadSleepTime);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Capture thread error: {ex.Message}");
            }
        }

        private void InitializeTracking()
        {
            // Connect to AR glasses using the head tracking manager
            if (headTracking.Initialize())
            {
                // Start update loops
                updateTimer.Start();
                if (!config.UseThreadedCapture && captureTimer != null)
                    captureTimer.Start();
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

        private int GetSourceDisplayIndex()
        {
            // Convert relative indices to absolute
            int displayCount = Screen.AllScreens.Length;
            int actualDisplayIndex = config.SourceDisplayIndex;
            
            if (config.SourceDisplayIndex < 0)
            {
                actualDisplayIndex = displayCount + config.SourceDisplayIndex;
                
                // Make sure we don't go below 0
                if (actualDisplayIndex < 0)
                    actualDisplayIndex = 0;
            }
            else if (config.SourceDisplayIndex >= displayCount)
            {
                actualDisplayIndex = displayCount - 1;
            }
            
            return actualDisplayIndex;
        }

        private void UpdateSourceDisplayScaleFactor()
        {
            try
            {
                Screen sourceScreen = Screen.AllScreens[GetSourceDisplayIndex()];
                IntPtr hwnd = MonitorFromWindow(IntPtr.Zero, MONITOR_DEFAULTTONEAREST);
                
                if (GetDpiForMonitor(hwnd, MDT_EFFECTIVE_DPI, out uint dpiX, out uint dpiY))
                {
                    // Standard DPI is 96
                    sourceScaleFactor = dpiX / 96.0f;
                    Console.WriteLine($"Source display scale factor: {sourceScaleFactor}");
                }
                else
                {
                    // Default to system DPI
                    using (Graphics g = Graphics.FromHwnd(IntPtr.Zero))
                    {
                        sourceScaleFactor = g.DpiX / 96.0f;
                        Console.WriteLine($"Source display scale factor (fallback): {sourceScaleFactor}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error detecting display scale factor: {ex.Message}");
                sourceScaleFactor = 1.0f; // Default to no scaling
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

        private float ToDegrees(float radians)
        {
            return radians * 180.0f / (float)Math.PI;
        }

        private void CaptureTimer_Tick(object sender, EventArgs e)
        {
            var stopWatch = System.Diagnostics.Stopwatch.StartNew();

            // Capture the screen from the source display
            CaptureScreen();
            
            // Scale and update the display
            UpdateContentImage();

            // Print the time taken in milliseconds for debugging
            Console.WriteLine($"CaptureTimer_Tick time: {stopWatch.ElapsedMilliseconds} ms");
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            // Signal threads to stop
            shutdownThreads = true;
            
            // Stop timers
            updateTimer.Stop();
            if (!config.UseThreadedCapture && captureTimer != null)
                captureTimer.Stop();
            
            // Wait for thread to exit
            if (captureThread != null && captureThread.IsAlive)
            {
                try
                {
                    captureThread.Join(500); // Wait up to 500ms for thread to exit
                }
                catch { }
            }
            
            // Disconnect from AR glasses
            headTracking.Shutdown();
            
            // Clean up resources
            lock (this)
            {
                capturedScreen?.Dispose();
                scaledCapture?.Dispose();
            }
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

        private void CaptureScreen()
        {
            try
            {
                // Get the source screen to capture
                Screen sourceScreen = Screen.AllScreens[GetSourceDisplayIndex()];
                Rectangle screenBounds = sourceScreen.Bounds;
                
                // Determine the actual area to capture
                Rectangle captureArea = captureRegion.IsEmpty ? screenBounds : captureRegion;
                
                // Apply scaling factor to capture correct physical pixel resolution
                int actualWidth = (int)(captureArea.Width * sourceScaleFactor);
                int actualHeight = (int)(captureArea.Height * sourceScaleFactor);
                int actualLeft = (int)(captureArea.Left * sourceScaleFactor);
                int actualTop = (int)(captureArea.Top * sourceScaleFactor);
                
                if (config.UseDirectCapture)
                {
                    // Direct capture using BitBlt for better performance
                    IntPtr desktopDC = GetDC(IntPtr.Zero);
                    
                    try
                    {
                        // Create or reuse bitmap
                        if (capturedScreen == null || capturedScreen.Width != actualWidth || capturedScreen.Height != actualHeight)
                        {
                            lock (this) // Prevent race condition with UpdateContentImage
                            {
                                capturedScreen?.Dispose();
                                capturedScreen = new Bitmap(actualWidth, actualHeight, PixelFormat.Format32bppRgb);
                            }
                        }
                        
                        using (Graphics g = Graphics.FromImage(capturedScreen))
                        {
                            // Set up graphics for speed
                            g.CompositingMode = CompositingMode.SourceCopy; // Fastest
                            g.CompositingQuality = CompositingQuality.HighSpeed;
                            g.InterpolationMode = InterpolationMode.NearestNeighbor;
                            g.SmoothingMode = SmoothingMode.None;
                            g.PixelOffsetMode = PixelOffsetMode.None;
                            
                            IntPtr hdc = g.GetHdc();
                            try
                            {
                                // Copy screen to bitmap in one operation
                                BitBlt(hdc, 0, 0, actualWidth, actualHeight,
                                      desktopDC, actualLeft, actualTop, SRCCOPY);
                            }
                            finally
                            {
                                g.ReleaseHdc(hdc);
                            }
                        }
                        
                        DrawCursorOnBitmap(capturedScreen, captureArea);
                    }
                    finally
                    {
                        ReleaseDC(IntPtr.Zero, desktopDC);
                    }
                }
                else
                {
                    // Original capture method as fallback
                    IntPtr desktopDC = GetDC(IntPtr.Zero);
                    IntPtr memoryDC = CreateCompatibleDC(desktopDC);
                    IntPtr captureBitmap = CreateCompatibleBitmap(desktopDC, actualWidth, actualHeight);
                    IntPtr oldObject = SelectObject(memoryDC, captureBitmap);
                    
                    bool success = BitBlt(
                        memoryDC, 0, 0, actualWidth, actualHeight,
                        desktopDC, actualLeft, actualTop, SRCCOPY);
                    
                    if (success)
                    {
                        if (capturedScreen != null)
                        {
                            capturedScreen.Dispose();
                        }
                        
                        capturedScreen = Bitmap.FromHbitmap(captureBitmap);
                        DrawCursorOnBitmap(capturedScreen, captureArea);
                    }
                    
                    SelectObject(memoryDC, oldObject);
                    DeleteObject(captureBitmap);
                    DeleteDC(memoryDC);
                    ReleaseDC(IntPtr.Zero, desktopDC);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error capturing screen: {ex.Message}");
            }
        }

        private void DrawCursorOnBitmap(Bitmap bitmap, Rectangle screenBounds)
        {
            // Get cursor information
            CURSORINFO cursorInfo = new CURSORINFO();
            cursorInfo.cbSize = Marshal.SizeOf(cursorInfo);
            
            if (GetCursorInfo(out cursorInfo) && (cursorInfo.flags & CURSOR_SHOWING) != 0)
            {
                // Get cursor position and handle
                IntPtr cursorHandle = cursorInfo.hCursor;
                POINT cursorPosition = cursorInfo.ptScreenPos;
                
                // Check if cursor is in capture area
                if (cursorPosition.x >= screenBounds.Left && cursorPosition.x <= screenBounds.Right &&
                    cursorPosition.y >= screenBounds.Top && cursorPosition.y <= screenBounds.Bottom)
                {
                    // Get hotspot information to properly position the cursor
                    ICONINFO iconInfo;
                    if (GetIconInfo(cursorHandle, out iconInfo))
                    {
                        try
                        {
                            // Calculate cursor position relative to screen bounds, accounting for scaling
                            int cursorX = (int)((cursorPosition.x - screenBounds.Left) * sourceScaleFactor) - iconInfo.xHotspot;
                            int cursorY = (int)((cursorPosition.y - screenBounds.Top) * sourceScaleFactor) - iconInfo.yHotspot;
                            
                            // Draw cursor on the bitmap
                            using (Graphics g = Graphics.FromImage(bitmap))
                            {
                                IntPtr hdc = g.GetHdc();
                                DrawIcon(hdc, cursorX, cursorY, cursorHandle);
                                g.ReleaseHdc(hdc);
                            }
                            
                            // Clean up icon resources
                            if (iconInfo.hbmColor != IntPtr.Zero)
                                DeleteObject(iconInfo.hbmColor);
                            if (iconInfo.hbmMask != IntPtr.Zero)
                                DeleteObject(iconInfo.hbmMask);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error drawing cursor: {ex.Message}");
                        }
                    }
                }
            }
        }

        private void UpdateContentImage()
        {
            try
            {
                if (capturedScreen != null)
                {
                    var resizeStopwatch = System.Diagnostics.Stopwatch.StartNew();
                    
                    // Only create a new bitmap if necessary
                    lock (this) // Prevent race condition with CaptureScreen
                    {
                        if (scaledCapture == null || 
                            scaledCapture.Width != config.ContentSize.Width || 
                            scaledCapture.Height != config.ContentSize.Height)
                        {
                            scaledCapture?.Dispose();
                            scaledCapture = new Bitmap(config.ContentSize.Width, config.ContentSize.Height, PixelFormat.Format32bppRgb);
                        }
                        
                        // Calculate scaled dimensions to maintain aspect ratio
                        Rectangle sourceRect = new Rectangle(0, 0, capturedScreen.Width, capturedScreen.Height);
                        Rectangle destRect = CalculateScaledRectangle(sourceRect, config.ContentSize);
                        
                        // Use optimized scaling based on performance settings
                        using (Graphics g = Graphics.FromImage(scaledCapture))
                        {
                            // Set quality based on performance settings
                            if (config.UseLowQualityScaling)
                            {
                                g.CompositingMode = CompositingMode.SourceCopy; // Fastest
                                g.InterpolationMode = InterpolationMode.Default;
                                g.SmoothingMode = SmoothingMode.Default;
                                g.CompositingQuality = CompositingQuality.HighSpeed;
                                g.PixelOffsetMode = PixelOffsetMode.None;
                            }
                            else
                            {
                                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                                g.SmoothingMode = SmoothingMode.HighQuality;
                                g.CompositingQuality = CompositingQuality.HighQuality;
                                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                            }
                            
                            // Clear background only if needed (source doesn't fill the entire area)
                            if (destRect.Width < config.ContentSize.Width || destRect.Height < config.ContentSize.Height)
                            {
                                g.Clear(Color.Black);
                            }
                            
                            // Draw image
                            g.DrawImage(capturedScreen, destRect, sourceRect, GraphicsUnit.Pixel);
                        }
                    }
                    
                    //Console.WriteLine($"Image scaling time: {resizeStopwatch.ElapsedMilliseconds} ms");
                    
                    // Update UI on the UI thread if needed
                    if (InvokeRequired)
                    {
                        BeginInvoke(new Action(() => {
                            // Set image without creating a new reference (avoids unnecessary redraws)
                            if (contentDisplay.Image == scaledCapture)
                            {
                                contentDisplay.Invalidate(); // Lighter than Refresh()
                            }
                            else
                            {
                                contentDisplay.Image = scaledCapture;
                            }
                        }));
                    }
                    else
                    {
                        // Set image without creating a new reference (avoids unnecessary redraws)
                        if (contentDisplay.Image == scaledCapture)
                        {
                            contentDisplay.Invalidate(); // Lighter than Refresh()
                        }
                        else
                        {
                            contentDisplay.Image = scaledCapture;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error updating content image: {ex.Message}");
            }
        }
        
        // Helper method to calculate the destination rectangle for proper aspect ratio scaling
        private Rectangle CalculateScaledRectangle(Rectangle srcRect, Size destSize)
        {
            // Calculate source aspect ratio
            float srcAspectRatio = (float)srcRect.Width / srcRect.Height;
            
            // Calculate destination aspect ratio
            float destAspectRatio = (float)destSize.Width / destSize.Height;
            
            // Determine dimensions based on aspect ratios
            int destWidth, destHeight;
            int destX = 0, destY = 0;
            
            if (srcAspectRatio >= destAspectRatio)
            {
                // Source is wider than destination, fit to width
                destWidth = destSize.Width;
                destHeight = (int)(destWidth / srcAspectRatio);
                destY = (destSize.Height - destHeight) / 2; // Center vertically
            }
            else
            {
                // Source is taller than destination, fit to height
                destHeight = destSize.Height;
                destWidth = (int)(destHeight * srcAspectRatio);
                destX = (destSize.Width - destWidth) / 2; // Center horizontally
            }
            
            return new Rectangle(destX, destY, destWidth, destHeight);
        }
        
        // Helper method for cursor position scaling
        private Point ScaleCursorPosition(Point cursorPos, Rectangle sourceBounds, Rectangle targetRect)
        {
            // Calculate the scaling factors
            float scaleX = (float)targetRect.Width / sourceBounds.Width;
            float scaleY = (float)targetRect.Height / sourceBounds.Height;
            
            // Calculate the scaled position relative to the target rectangle
            int scaledX = targetRect.X + (int)((cursorPos.X - sourceBounds.X) * scaleX);
            int scaledY = targetRect.Y + (int)((cursorPos.Y - sourceBounds.Y) * scaleY);
            
            return new Point(scaledX, scaledY);
        }

        // Optional method to set capture region (e.g., for capturing just a window instead of full screen)
        public void SetCaptureRegion(Rectangle region)
        {
            captureRegion = region;
        }

        // Optional method to reset to full screen capture
        public void ResetCaptureRegion()
        {
            captureRegion = Rectangle.Empty;
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