using System;
using System.Drawing;
using System.Numerics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Drawing.Imaging;
using System.Drawing.Drawing2D;
using Timer = System.Windows.Forms.Timer;

namespace ARContentStabilizer
{
    public class ScreenCaptureManager
    {
        #region DLL Imports

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

        // Constants for cursor
        private const int CURSOR_SHOWING = 0x00000001;

        // Constants for BitBlt
        private const int SRCCOPY = 0xCC0020;

        // Constants for DPI awareness
        private const uint MONITOR_DEFAULTTONEAREST = 0x00000002;
        private const int MDT_EFFECTIVE_DPI = 0;

        #endregion

        #region Properties and Fields

        // Configuration reference
        private readonly ConfigManager config;
        
        // UI Components
        private PictureBox contentDisplay;
        
        // Bitmaps for screen capture
        private Bitmap capturedScreen;
        private Bitmap scaledCapture;
        
        // Scaling factors
        private float sourceScaleFactor = 1.0f;

        // Performance settings
        private Rectangle captureRegion = Rectangle.Empty; // Optional region to capture instead of full screen
        private bool captureInProgress = false;  // Flag to prevent overlapping captures
        private System.Threading.Thread captureThread; // Background thread for capture
        private volatile bool shutdownThreads = false; // Signal to shutdown background threads
        
        // Timer for capture
        private Timer captureTimer;

        #endregion

        #region Constructor and Initialization

        public ScreenCaptureManager(ConfigManager config, PictureBox displayControl)
        {
            this.config = config;
            this.contentDisplay = displayControl;
            
            // Create initial bitmaps
            capturedScreen = new Bitmap(Screen.AllScreens[GetSourceDisplayIndex()].Bounds.Width, 
                                      Screen.AllScreens[GetSourceDisplayIndex()].Bounds.Height);
            scaledCapture = new Bitmap(config.ContentSize.Width, config.ContentSize.Height);
            
            // Initialize display control
            contentDisplay.Image = scaledCapture;
            
            // Detect source display scaling factor
            UpdateSourceDisplayScaleFactor();
            
            // Set up capture mechanism
            SetupCapture();
        }

        private void SetupCapture()
        {
            if (config.UseThreadedCapture)
            {
                // Use a dedicated thread for capturing at maximum possible rate
                captureThread = new System.Threading.Thread(CaptureThreadMethod);
                captureThread.IsBackground = true;
                captureThread.Priority = System.Threading.ThreadPriority.AboveNormal;
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

        #region Capture and Display Methods

        public void Start()
        {
            if (config.UseThreadedCapture)
            {
                captureThread.Start();
            }
            else if (captureTimer != null)
            {
                captureTimer.Start();
            }
        }

        public void Stop()
        {
            // Signal threads to stop
            shutdownThreads = true;
            
            // Stop timer
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
            
            // Clean up resources
            lock (this)
            {
                capturedScreen?.Dispose();
                scaledCapture?.Dispose();
            }
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
                    if (contentDisplay.InvokeRequired)
                    {
                        contentDisplay.BeginInvoke(new Action(() => {
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

        #endregion

        #region Utility Methods

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

        // Set a specific region to capture
        public void SetCaptureRegion(Rectangle region)
        {
            captureRegion = region;
        }

        // Reset to capture the full screen
        public void ResetCaptureRegion()
        {
            captureRegion = Rectangle.Empty;
        }

        #endregion
    }
} 