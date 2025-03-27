using System;
using System.Drawing;
using System.Numerics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Drawing.Imaging;
using System.Drawing.Drawing2D;
using SharpDX;
using SharpDX.Direct3D11;
using SharpDX.DXGI;
using Device = SharpDX.Direct3D11.Device;
using Resource = SharpDX.DXGI.Resource;
using Timer = System.Windows.Forms.Timer;

namespace ARContentStabilizer
{
    public class ScreenCaptureManager
    {
        #region DLL Imports

        // DPI awareness imports
        [DllImport("user32.dll")]
        static extern IntPtr MonitorFromWindow(IntPtr hwnd, uint dwFlags);

        [DllImport("user32.dll")]
        static extern bool GetDpiForMonitor(IntPtr hmonitor, int dpiType, out uint dpiX, out uint dpiY);

        // Memory copying import
        [DllImport("kernel32.dll", EntryPoint = "RtlMoveMemory", SetLastError = false)]
        private static extern void CopyMemory(IntPtr dest, IntPtr src, int count);

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

        // Desktop Duplication API objects
        private Factory1 factory;
        private Adapter1 adapter;
        private Device device;
        private DeviceContext context;
        private OutputDuplication outputDuplication;
        private Output1 output1;
        private Texture2D stagingTexture;
        private int adapterOutput;

        // Add these structures at the top of the class, in the Properties and Fields region
        private struct CursorInfo
        {
            public bool IsVisible;
            public Point Position;
            public SharpDX.DXGI.OutputDuplicatePointerPosition PointerPosition;
            public OutputDuplicatePointerShapeInformation ShapeInfo;
            public byte[] ShapeBuffer;
        }

        private CursorInfo currentCursor;

        // Add these enums in the Properties and Fields region
        private enum CursorShapeType
        {
            MonoChrome = 1,
            Color = 2,
            MaskedColor = 4
        }

        #endregion

        #region Constructor and Initialization

        public ScreenCaptureManager(ConfigManager config, PictureBox displayControl)
        {
            this.config = config;
            this.contentDisplay = displayControl;
            
            // Create initial bitmaps
            scaledCapture = new Bitmap(config.ContentSize.Width, config.ContentSize.Height);
            
            // Initialize display control
            contentDisplay.Image = scaledCapture;
            
            // Detect source display scaling factor
            UpdateSourceDisplayScaleFactor();
            
            // Initialize Desktop Duplication API
            InitializeDesktopDuplication();
            
            // Set up capture mechanism
            SetupCapture();
        }

        private void InitializeDesktopDuplication()
        {
            try
            {
                // Create DXGI Factory
                factory = new Factory1();
                
                int adapterCount = factory.GetAdapterCount1();
                Console.WriteLine($"Found {adapterCount} adapters");
                
                // Try each adapter
                for (int i = 0; i < adapterCount; i++)
                {
                    try
                    {
                        adapter = factory.GetAdapter1(i);
                        Console.WriteLine($"Trying adapter {i}: {adapter.Description.Description}");
                        
                        // Continue with device creation...
                        // If successful, break out of the loop
                        break;
                    }
                    catch
                    {
                        // Clean up and try next adapter
                        adapter?.Dispose();
                        adapter = null;
                    }
                }
                
                if (adapter == null)
                {
                    throw new Exception("No suitable graphics adapter found");
                }
                
                // Create Direct3D device
                device = new Device(adapter, DeviceCreationFlags.BgraSupport);
                Console.WriteLine("D3D device created successfully");
                
                // Create device context
                context = device.ImmediateContext;
                
                // Get source display index
                adapterOutput = GetSourceDisplayIndex();
                
                try
                {
                    // Find the right output (monitor)
                    var output = adapter.GetOutput(adapterOutput);
                    Console.WriteLine($"Using display: {output.Description.DeviceName}");
                    
                    output1 = output.QueryInterface<Output1>();
                    output.Dispose();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to get output: {ex.Message}");
                    throw;
                }
                
                try
                {
                    // Create output duplication for this output
                    outputDuplication = output1.DuplicateOutput(device);
                    Console.WriteLine("Output duplication created successfully");
                }
                catch (SharpDXException ex)
                {
                    // Check specific DXGI error codes
                    if (ex.ResultCode.Code == SharpDX.DXGI.ResultCode.NotCurrentlyAvailable.Result.Code)
                    {
                        Console.WriteLine("Desktop Duplication is already in use by another application");
                        throw new InvalidOperationException("Desktop Duplication is already in use by another application");
                    }
                    else if (ex.ResultCode.Code == SharpDX.DXGI.ResultCode.Unsupported.Result.Code)
                    {
                        Console.WriteLine("Desktop Duplication is not supported on this system");
                        throw new NotSupportedException("Desktop Duplication is not supported on this system");
                    }
                    else
                    {
                        Console.WriteLine($"Failed to create output duplication: {ex.Message}, Code: {ex.ResultCode}");
                        throw;
                    }
                }
                
                // Get output description to determine screen size
                var outputDesc = output1.Description;
                Rectangle bounds = new Rectangle(outputDesc.DesktopBounds.Left, 
                                                outputDesc.DesktopBounds.Top,
                                                outputDesc.DesktopBounds.Right - outputDesc.DesktopBounds.Left,
                                                outputDesc.DesktopBounds.Bottom - outputDesc.DesktopBounds.Top);
                
                Console.WriteLine($"Capture bounds: {bounds.Width}x{bounds.Height} at {bounds.X},{bounds.Y}");
                
                // Create staging texture for CPU access
                var textureDesc = new Texture2DDescription
                {
                    CpuAccessFlags = CpuAccessFlags.Read,
                    BindFlags = BindFlags.None,
                    Format = Format.B8G8R8A8_UNorm,
                    Width = bounds.Width,
                    Height = bounds.Height,
                    OptionFlags = ResourceOptionFlags.None,
                    MipLevels = 1,
                    ArraySize = 1,
                    SampleDescription = { Count = 1, Quality = 0 },
                    Usage = ResourceUsage.Staging
                };
                
                stagingTexture = new Texture2D(device, textureDesc);
                
                // Create the bitmap with screen dimensions
                capturedScreen = new Bitmap(bounds.Width, bounds.Height, PixelFormat.Format32bppArgb);
                
                Console.WriteLine($"Desktop Duplication API initialized for output {adapterOutput} with dimensions {bounds.Width}x{bounds.Height}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error initializing Desktop Duplication API: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
                // Clean up any resources that were created
                CleanupDuplicationResources();
                throw;
            }
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
            int displayCount = adapter.GetOutputCount();
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
            
            Console.WriteLine($"Display selection: requested={config.SourceDisplayIndex}, actual={actualDisplayIndex}, total={displayCount}");
            
            // Also log information about all available displays
            for (int i = 0; i < displayCount; i++)
            {
                using (var tmpOutput = adapter.GetOutput(i))
                {
                    Console.WriteLine($"Display {i}: {tmpOutput.Description.DeviceName}, " +
                                     $"Bounds: {tmpOutput.Description.DesktopBounds.Left},{tmpOutput.Description.DesktopBounds.Top} " +
                                     $"{tmpOutput.Description.DesktopBounds.Right-tmpOutput.Description.DesktopBounds.Left}x" +
                                     $"{tmpOutput.Description.DesktopBounds.Bottom-tmpOutput.Description.DesktopBounds.Top}");
                }
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
                
                // Clean up DirectX resources
                CleanupDuplicationResources();
            }
        }

        private void CleanupDuplicationResources()
        {
            // Dispose of all DirectX resources
            stagingTexture?.Dispose();
            stagingTexture = null;
            
            outputDuplication?.Dispose();
            outputDuplication = null;
            
            output1?.Dispose();
            output1 = null;
            
            context?.Dispose();
            context = null;
            
            device?.Dispose();
            device = null;
            
            adapter?.Dispose();
            adapter = null;
            
            factory?.Dispose();
            factory = null;
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
            // Check if testing is enabled
            if (config.EnableTestPattern)
            {
                DrawTestPattern();
                return;
            }

            // Check if Desktop Duplication API is initialized
            if (outputDuplication == null)
            {
                Console.WriteLine("Desktop Duplication API not initialized, using fallback capture");
                FallbackCaptureScreen();
                return;
            }

            // Try to get the next frame
            SharpDX.DXGI.Resource desktopResource = null;
            OutputDuplicateFrameInformation frameInfo;
            Result result = Result.Ok; // Define result outside the try block

            try
            {
                // Wait to get the next frame
                result = outputDuplication.TryAcquireNextFrame(500, out frameInfo, out desktopResource);
                
                // Log frame information for debugging
                Console.WriteLine($"Frame acquired: {result.Success}, AccumulatedFrames: {frameInfo.AccumulatedFrames}");
                
                if (desktopResource != null && result.Success)
                {
                    try
                    {
                        // Update cursor information
                        UpdateCursor(frameInfo);

                        // Get the desktop image
                        using (var desktopTexture = desktopResource.QueryInterface<Texture2D>())
                        {
                            // Log texture details
                            var desc = desktopTexture.Description;
                            Console.WriteLine($"Texture: {desc.Width}x{desc.Height}, Format={desc.Format}, " +
                                             $"Usage={desc.Usage}, Samples={desc.SampleDescription.Count}");
                            
                            // Copy the desktop texture to a staging texture that allows for CPU access
                            context.CopyResource(desktopTexture, stagingTexture);
                            
                            // Log when copy is complete
                            Console.WriteLine("Resource copied to staging texture");
                            
                            // Map the staging texture to get access to the pixel data
                            var dataBox = context.MapSubresource(stagingTexture, 0, MapMode.Read, SharpDX.Direct3D11.MapFlags.None);
                            
                            try
                            {
                                // Create a lock around bitmap operations to ensure thread safety
                                lock (capturedScreen)
                                {
                                    // Copy the pixel data to the bitmap
                                    var rect = new Rectangle(0, 0, capturedScreen.Width, capturedScreen.Height);
                                    BitmapData bitmapData = null;
                                    
                                    try
                                    {
                                        bitmapData = capturedScreen.LockBits(rect, ImageLockMode.WriteOnly, capturedScreen.PixelFormat);
                                        
                                        // Make sure we're copying the correct number of bytes
                                        int bytesToCopy = Math.Min(dataBox.RowPitch, bitmapData.Stride);
                                        
                                        // Copy each line (need to handle stride properly)
                                        for (int y = 0; y < capturedScreen.Height; y++)
                                        {
                                            IntPtr sourcePtr = dataBox.DataPointer + y * dataBox.RowPitch;
                                            IntPtr destPtr = bitmapData.Scan0 + y * bitmapData.Stride;
                                            CopyMemory(destPtr, sourcePtr, bytesToCopy);
                                        }
                                    }
                                    finally
                                    {
                                        // Always unlock the bits, even if an error occurred
                                        if (bitmapData != null)
                                        {
                                            capturedScreen.UnlockBits(bitmapData);
                                        }
                                    }
                                }
                            }
                            finally
                            {
                                // Always unmap the texture, even if an error occurred
                                context.UnmapSubresource(stagingTexture, 0);
                            }
                        }

                        // After copying the screen content, draw the cursor
                        if (currentCursor.IsVisible && currentCursor.ShapeBuffer != null)
                        {
                            string cursorTypeStr = "Unknown";
                            switch (currentCursor.ShapeInfo.Type)
                            {
                                case (int)CursorShapeType.MonoChrome: cursorTypeStr = "MonoChrome"; break;
                                case (int)CursorShapeType.Color: cursorTypeStr = "Color"; break;
                                case (int)CursorShapeType.MaskedColor: cursorTypeStr = "MaskedColor"; break;
                            }
                            
                            Console.WriteLine($"Drawing cursor: Type={cursorTypeStr}, Size={currentCursor.ShapeInfo.Width}x{currentCursor.ShapeInfo.Height}, " +
                                             $"Position={currentCursor.Position.X},{currentCursor.Position.Y}");

                            lock (capturedScreen)
                            {
                                using (Graphics g = Graphics.FromImage(capturedScreen))
                                {
                                    if (currentCursor.ShapeInfo.Type == (int)CursorShapeType.Color)
                                    {
                                        // Handle color cursor
                                        using (var cursorBitmap = new Bitmap(
                                            currentCursor.ShapeInfo.Width,
                                            currentCursor.ShapeInfo.Height,
                                            PixelFormat.Format32bppArgb))
                                        {
                                            var bitmapData = cursorBitmap.LockBits(
                                                new Rectangle(0, 0, cursorBitmap.Width, cursorBitmap.Height),
                                                ImageLockMode.WriteOnly,
                                                PixelFormat.Format32bppArgb);

                                            Marshal.Copy(currentCursor.ShapeBuffer, 0, bitmapData.Scan0,
                                                Math.Min(currentCursor.ShapeBuffer.Length, bitmapData.Stride * bitmapData.Height));

                                            cursorBitmap.UnlockBits(bitmapData);

                                            g.DrawImage(cursorBitmap,
                                                currentCursor.Position.X,
                                                currentCursor.Position.Y);
                                        }
                                    }
                                    else if (currentCursor.ShapeInfo.Type == (int)CursorShapeType.MaskedColor)
                                    {
                                        // Handle masked color cursor
                                        using (var cursorBitmap = new Bitmap(
                                            currentCursor.ShapeInfo.Width,
                                            currentCursor.ShapeInfo.Height,
                                            PixelFormat.Format32bppArgb))
                                        {
                                            var bitmapData = cursorBitmap.LockBits(
                                                new Rectangle(0, 0, cursorBitmap.Width, cursorBitmap.Height),
                                                ImageLockMode.WriteOnly,
                                                PixelFormat.Format32bppArgb);

                                            try
                                            {
                                                int stride = currentCursor.ShapeInfo.Pitch;
                                                int height = currentCursor.ShapeInfo.Height;
                                                int width = currentCursor.ShapeInfo.Width;
                                                
                                                // First copy the XOR mask (color data)
                                                Marshal.Copy(currentCursor.ShapeBuffer, 0, bitmapData.Scan0, 
                                                    Math.Min(stride * height, bitmapData.Stride * height));

                                                // Verify we have enough buffer for the AND mask
                                                int maskOffset = stride * height;
                                                if (currentCursor.ShapeBuffer.Length > maskOffset)
                                                {
                                                    // Rewrite without unsafe code
                                                    for (int y = 0; y < height; y++)
                                                    {
                                                        for (int x = 0; x < width; x++)
                                                        {
                                                            int byteIndex = maskOffset + (y * stride) + (x / 8);
                                                            
                                                            // Add bounds checking to prevent array access out of bounds
                                                            if (byteIndex >= currentCursor.ShapeBuffer.Length)
                                                                continue;
                                                                
                                                            int bitIndex = 7 - (x % 8);
                                                            bool andBit = (currentCursor.ShapeBuffer[byteIndex] & (1 << bitIndex)) != 0;
                                                            
                                                            if (andBit)
                                                            {
                                                                // Get the offset to the pixel in the bitmap data
                                                                int pixelOffset = (y * bitmapData.Stride) + (x * 4);
                                                                
                                                                // Set alpha to 0 (transparent) while preserving RGB
                                                                Marshal.WriteByte(bitmapData.Scan0 + pixelOffset + 3, 0);
                                                            }
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    Console.WriteLine($"Cursor buffer too small for AND mask. Expected offset: {maskOffset}, buffer length: {currentCursor.ShapeBuffer.Length}");
                                                }
                                            }
                                            finally
                                            {
                                                cursorBitmap.UnlockBits(bitmapData);
                                            }

                                            // Enable high-quality rendering for the cursor
                                            g.InterpolationMode = InterpolationMode.NearestNeighbor;
                                            g.CompositingMode = CompositingMode.SourceOver;
                                            g.CompositingQuality = CompositingQuality.HighQuality;
                                            g.SmoothingMode = SmoothingMode.None;
                                            g.PixelOffsetMode = PixelOffsetMode.Half;

                                            g.DrawImage(cursorBitmap,
                                                currentCursor.Position.X,
                                                currentCursor.Position.Y);
                                        }
                                    }
                                    else if (currentCursor.ShapeInfo.Type == (int)CursorShapeType.MonoChrome && 
                                        currentCursor.ShapeInfo.Width < 10) // Typical text cursor is very narrow
                                    {
                                        // Create a simpler, better-looking text cursor
                                        using (var cursorBitmap = new Bitmap(
                                            currentCursor.ShapeInfo.Width,
                                            currentCursor.ShapeInfo.Height,
                                            PixelFormat.Format32bppArgb))
                                        {
                                            using (Graphics cursorG = Graphics.FromImage(cursorBitmap))
                                            {
                                                // Draw a simple vertical line for text cursor
                                                cursorG.Clear(Color.Transparent);
                                                using (Pen cursorPen = new Pen(Color.White, 1))
                                                {
                                                    int middle = cursorBitmap.Width / 2;
                                                    cursorG.DrawLine(cursorPen, 
                                                        middle, 0,
                                                        middle, cursorBitmap.Height);
                                                }
                                            }
                                            
                                            // Draw it at the right position
                                            g.DrawImage(cursorBitmap,
                                                currentCursor.Position.X,
                                                currentCursor.Position.Y);
                                        }
                                    }
                                    else
                                    {
                                        try
                                        {
                                            // Handle monochrome cursor (like text cursor)
                                            int width = currentCursor.ShapeInfo.Width;
                                            int height = currentCursor.ShapeInfo.Height;
                                            int pitch = currentCursor.ShapeInfo.Pitch;

                                            // Add validation
                                            if (width <= 0 || height <= 0 || pitch <= 0 || currentCursor.ShapeBuffer == null)
                                            {
                                                Console.WriteLine("Invalid cursor dimensions or null buffer");
                                                return;
                                            }

                                            using (var cursorBitmap = new Bitmap(width, height, PixelFormat.Format32bppArgb))
                                            {
                                                // Process XOR and AND masks
                                                int maskSize = (pitch * height);
                                                
                                                // Validate buffer size before proceeding
                                                if (currentCursor.ShapeBuffer.Length < maskSize * 2)
                                                {
                                                    Console.WriteLine($"Cursor buffer too small. Expected {maskSize * 2}, got {currentCursor.ShapeBuffer.Length}");
                                                    return;
                                                }

                                                byte[] andMask = new byte[maskSize];
                                                byte[] xorMask = new byte[maskSize];

                                                try
                                                {
                                                    Array.Copy(currentCursor.ShapeBuffer, 0, andMask, 0, maskSize);
                                                    Array.Copy(currentCursor.ShapeBuffer, maskSize, xorMask, 0, maskSize);

                                                    // Convert monochrome masks to actual pixels
                                                    for (int y = 0; y < height; y++)
                                                    {
                                                        for (int x = 0; x < width; x++)
                                                        {
                                                            // Calculate indices with bounds checking
                                                            int byteIndex = (y * pitch) + (x / 8);
                                                            if (byteIndex >= maskSize)
                                                            {
                                                                Console.WriteLine($"Invalid byte index: {byteIndex}, maskSize: {maskSize}");
                                                                continue;
                                                            }

                                                            int bitIndex = 7 - (x % 8);
                                                            
                                                            bool andBit = (andMask[byteIndex] & (1 << bitIndex)) != 0;
                                                            bool xorBit = (xorMask[byteIndex] & (1 << bitIndex)) != 0;

                                                            Color pixelColor;
                                                            
                                                            // FIXED: Correct logic for monochrome cursor rendering
                                                            // Windows monochrome cursors use the AND mask to create a screen hole
                                                            // and the XOR mask to draw the cursor itself (inverted color)
                                                            if (andBit)
                                                            {
                                                                // AND bit is 1 - screen mask (to clear background)
                                                                if (xorBit)
                                                                    // Where AND=1 and XOR=1: Draw inverse screen (for text cursor)
                                                                    pixelColor = Color.FromArgb(128, 255, 255, 255); // Semi-transparent white (inverting effect)
                                                                else
                                                                    // Where AND=1 and XOR=0: Draw transparent black (screen hole)
                                                                    pixelColor = Color.Transparent;
                                                            }
                                                            else
                                                            {
                                                                // AND bit is 0 - opaque cursor pixel
                                                                if (xorBit)
                                                                    // Where AND=0 and XOR=1: Draw white
                                                                    pixelColor = Color.White;
                                                                else
                                                                    // Where AND=0 and XOR=0: Draw transparent (not black)
                                                                    // This is the key fix for black background issues
                                                                    pixelColor = Color.Transparent;
                                                            }

                                                            cursorBitmap.SetPixel(x, y, pixelColor);
                                                        }
                                                    }

                                                    // Enable high-quality rendering for the cursor
                                                    g.InterpolationMode = InterpolationMode.NearestNeighbor;
                                                    g.CompositingMode = CompositingMode.SourceOver;
                                                    g.CompositingQuality = CompositingQuality.HighQuality;
                                                    g.SmoothingMode = SmoothingMode.None;
                                                    g.PixelOffsetMode = PixelOffsetMode.Half;

                                                    g.DrawImage(cursorBitmap,
                                                        currentCursor.Position.X,
                                                        currentCursor.Position.Y);
                                                }
                                                catch (Exception ex)
                                                {
                                                    Console.WriteLine($"Error processing cursor masks: {ex.Message}");
                                                    // Continue without drawing cursor
                                                }
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            Console.WriteLine($"Error drawing monochrome cursor: {ex.Message}");
                                            // Don't rethrow - allow capture to continue without cursor
                                        }
                                    }
                                }
                            }
                        }
                    }
                    finally
                    {
                        // Release the frame after we're done with it
                        outputDuplication.ReleaseFrame();
                        desktopResource?.Dispose();
                    }
                }
                else
                {
                    Console.WriteLine("No desktop resource acquired or frame acquisition failed");
                }
            }
            catch (SharpDXException ex) when (ex.ResultCode.Code == SharpDX.DXGI.ResultCode.WaitTimeout.Result.Code)
            {
                // Timeout is normal if no new frame is available
                Console.WriteLine("Timeout waiting for next frame");
            }
            catch (SharpDXException ex) when (ex.ResultCode.Code == SharpDX.DXGI.ResultCode.AccessLost.Result.Code)
            {
                // Access lost - desktop switch or user pressed Windows+L
                Console.WriteLine("Access to desktop lost. Attempting to reinitialize...");
                
                // Clean up and reinitialize
                CleanupDuplicationResources();
                InitializeDesktopDuplication();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error capturing screen: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
                
                // Draw test pattern as fallback when capture fails
                DrawTestPattern();
            }
        }

        private void UpdateContentImage()
        {
            try
            {
                if (capturedScreen != null)
                {
                    // Use local variables to prevent race conditions
                    Bitmap localCaptured;
                    
                    // Create a safe copy of the captured screen to avoid bitmap locking issues
                    lock (capturedScreen)
                    {
                        // Create a copy if needed
                        localCaptured = new Bitmap(capturedScreen);
                    }
                    
                    // Now process the copy without holding the lock
                    var resizeStopwatch = System.Diagnostics.Stopwatch.StartNew();
                    
                    // Create a new scaled capture if needed
                    lock (this)
                    {
                        if (scaledCapture == null || 
                            scaledCapture.Width != config.ContentSize.Width || 
                            scaledCapture.Height != config.ContentSize.Height)
                        {
                            scaledCapture?.Dispose();
                            scaledCapture = new Bitmap(config.ContentSize.Width, config.ContentSize.Height, PixelFormat.Format32bppRgb);
                        }
                        
                        // Calculate scaled dimensions to maintain aspect ratio
                        Rectangle sourceRect = new Rectangle(0, 0, localCaptured.Width, localCaptured.Height);
                        Rectangle destRect = CalculateScaledRectangle(sourceRect, config.ContentSize);
                        
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
                            
                            // Clear background if needed
                            if (destRect.Width < config.ContentSize.Width || destRect.Height < config.ContentSize.Height)
                            {
                                g.Clear(Color.Black);
                            }
                            
                            // Draw image
                            g.DrawImage(localCaptured, destRect, sourceRect, GraphicsUnit.Pixel);
                        }
                    }
                    
                    // Dispose of the local copy
                    localCaptured.Dispose();
                    
                    // Update UI on the UI thread
                    if (contentDisplay.InvokeRequired)
                    {
                        contentDisplay.BeginInvoke(new Action(() => {
                            if (contentDisplay.Image == scaledCapture)
                            {
                                contentDisplay.Invalidate();
                            }
                            else
                            {
                                contentDisplay.Image = scaledCapture;
                            }
                        }));
                    }
                    else
                    {
                        if (contentDisplay.Image == scaledCapture)
                        {
                            contentDisplay.Invalidate();
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
                Console.WriteLine(ex.StackTrace);
            }
        }

        private void FallbackCaptureScreen()
        {
            try
            {
                // Get screen bounds
                Screen screen = Screen.AllScreens[Math.Min(config.SourceDisplayIndex, Screen.AllScreens.Length - 1)];
                
                lock (capturedScreen)
                {
                    // Create a new bitmap if needed
                    if (capturedScreen == null || 
                        capturedScreen.Width != screen.Bounds.Width || 
                        capturedScreen.Height != screen.Bounds.Height)
                    {
                        capturedScreen?.Dispose();
                        capturedScreen = new Bitmap(screen.Bounds.Width, screen.Bounds.Height, PixelFormat.Format32bppArgb);
                    }
                    
                    using (Graphics g = Graphics.FromImage(capturedScreen))
                    {
                        // Capture the screen
                        g.CopyFromScreen(screen.Bounds.X, screen.Bounds.Y, 0, 0, screen.Bounds.Size);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Fallback capture failed: {ex.Message}");
                DrawTestPattern(); // Use test pattern as a last resort
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
            
            // If using Desktop Duplication API, we may need to reinitialize
            if (outputDuplication != null)
            {
                // For now, we'll ignore the region with Desktop Duplication as it's generally per-monitor
                Console.WriteLine("Note: Capture region is not fully supported with Desktop Duplication API.");
            }
        }

        // Reset to capture the full screen
        public void ResetCaptureRegion()
        {
            captureRegion = Rectangle.Empty;
        }

        private void DrawTestPattern()
        {
            // Draw a test pattern to verify the display pipeline
            lock (capturedScreen)
            {
                using (Graphics g = Graphics.FromImage(capturedScreen))
                {
                    // Clear to a color
                    g.Clear(Color.DarkBlue);
                    
                    // Draw some recognizable elements
                    using (Brush redBrush = new SolidBrush(Color.Red))
                    using (Brush greenBrush = new SolidBrush(Color.Green))
                    using (Pen whitePen = new Pen(Color.White, 3))
                    {
                        // Draw diagonal lines
                        g.DrawLine(whitePen, 0, 0, capturedScreen.Width, capturedScreen.Height);
                        g.DrawLine(whitePen, capturedScreen.Width, 0, 0, capturedScreen.Height);
                        
                        // Draw some shapes
                        g.FillRectangle(redBrush, capturedScreen.Width / 4, capturedScreen.Height / 4, 
                                       capturedScreen.Width / 2, capturedScreen.Height / 2);
                        g.FillEllipse(greenBrush, capturedScreen.Width / 3, capturedScreen.Height / 3, 
                                     capturedScreen.Width / 3, capturedScreen.Height / 3);
                    }
                    
                    // Draw text showing dimensions
                    using (Font font = new Font("Arial", 24))
                    using (Brush whiteBrush = new SolidBrush(Color.White))
                    {
                        g.DrawString($"Test Pattern {capturedScreen.Width}x{capturedScreen.Height}", 
                                    font, whiteBrush, 20, 20);
                        g.DrawString($"Source Display: {config.SourceDisplayIndex}", 
                                    font, whiteBrush, 20, 60);
                    }
                }
            }
        }

        // Add this method to update cursor information when capturing frames
        private void UpdateCursor(OutputDuplicateFrameInformation frameInfo)
        {
            try
            {
                if (frameInfo.LastMouseUpdateTime == 0)
                    return;

                // Update cursor position and visibility directly from frameInfo
                currentCursor.Position = new Point(
                    frameInfo.PointerPosition.Position.X,
                    frameInfo.PointerPosition.Position.Y
                );
                currentCursor.IsVisible = frameInfo.PointerPosition.Visible;

                // If the cursor shape has changed and has valid size
                if (frameInfo.PointerShapeBufferSize > 0)
                {
                    try
                    {
                        // Create unmanaged memory pointer for the shape buffer
                        IntPtr bufferPtr = Marshal.AllocHGlobal(frameInfo.PointerShapeBufferSize);
                        try
                        {
                            int bufferSize;
                            OutputDuplicatePointerShapeInformation shapeInfo;
                            
                            // Get the cursor shape
                            outputDuplication.GetFramePointerShape(
                                frameInfo.PointerShapeBufferSize,
                                bufferPtr,
                                out bufferSize,
                                out shapeInfo
                            );

                            // Validate the shape info and buffer size
                            if (bufferSize > 0 && 
                                shapeInfo.Width > 0 && 
                                shapeInfo.Height > 0 && 
                                shapeInfo.Pitch > 0)
                            {
                                // Calculate the expected buffer size based on shape info
                                int expectedSize = shapeInfo.Pitch * shapeInfo.Height;
                                if (shapeInfo.Type == (int)CursorShapeType.MonoChrome)
                                {
                                    expectedSize *= 2; // Monochrome cursors have AND and XOR masks
                                }

                                // Only update if the size makes sense
                                if (bufferSize <= frameInfo.PointerShapeBufferSize && 
                                    bufferSize >= expectedSize)
                                {
                                    // Allocate new buffer
                                    currentCursor.ShapeBuffer = new byte[bufferSize];
                                    currentCursor.ShapeInfo = shapeInfo;

                                    // Copy from unmanaged memory to managed buffer
                                    Marshal.Copy(bufferPtr, currentCursor.ShapeBuffer, 0, bufferSize);
                                }
                                else
                                {
                                    Console.WriteLine($"Invalid cursor buffer size: {bufferSize}, expected: {expectedSize}");
                                    currentCursor.ShapeBuffer = null;
                                }
                            }
                            else
                            {
                                Console.WriteLine("Invalid cursor shape info");
                                currentCursor.ShapeBuffer = null;
                            }
                        }
                        finally
                        {
                            Marshal.FreeHGlobal(bufferPtr);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error getting cursor shape: {ex.Message}");
                        currentCursor.ShapeBuffer = null;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error updating cursor: {ex.Message}");
                currentCursor.ShapeBuffer = null;
            }
        }

        #endregion
    }
} 