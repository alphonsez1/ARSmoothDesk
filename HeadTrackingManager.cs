using System;
using System.Numerics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Drawing;

namespace ARContentStabilizer
{
    public class HeadTrackingManager
    {
        #region DLL Imports

        // AirAPI DLL Imports
        [DllImport("deps\\AirAPI_Windows.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern int StartConnection();

        [DllImport("deps\\AirAPI_Windows.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern int StopConnection();

        [DllImport("deps\\AirAPI_Windows.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern IntPtr GetQuaternion();

        [DllImport("deps\\AirAPI_Windows.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern IntPtr GetEuler();

        #endregion

        #region Properties and Fields

        public Vector3 ViewFrontRay { get; private set; }
        public Vector3 ViewRightRay { get; private set; }
        public Vector3 ScreenCenterRay { get; private set; }
        public float Radius { get; set; }
        public bool IsConnected { get; private set; }

        // Position boundaries 
        private PointF minBoundary;
        private PointF maxBoundary;
        private PointF currentPosition;

        #endregion

        public HeadTrackingManager(float radius)
        {
            Radius = radius;
            IsConnected = false;
            minBoundary = new PointF(-100, -100); // Default values
            maxBoundary = new PointF(100, 100);   // Default values
            currentPosition = new PointF(0, 0);   // Start at center
        }

        public bool Initialize()
        {
            // Connect to AR glasses
            int connectionResult = StartConnection();
            if (connectionResult == 1)
            {
                Console.WriteLine("Connection to AR glasses started successfully");
                IsConnected = true;

                // Get initial rotation data
                UpdateRotationData();
                ScreenCenterRay = ViewFrontRay;
                Console.WriteLine($"Initial view center ray: {ViewFrontRay.X}, {ViewFrontRay.Y}, {ViewFrontRay.Z}");
                
                return true;
            }
            else
            {
                Console.WriteLine("Failed to connect to AR glasses");
                IsConnected = false;
                return false;
            }
        }

        public void Shutdown()
        {
            if (IsConnected)
            {
                StopConnection();
                IsConnected = false;
                Console.WriteLine("Connection to AR glasses stopped");
            }
        }

        public void UpdateRotationData()
        {
            // Get data from AirAPI
            var eulerPtr = GetEuler();
            // Order: roll, pitch, yaw
            var eulerArray = new float[3];
            Marshal.Copy(eulerPtr, eulerArray, 0, 3);

            Console.WriteLine("Euler angles: " +
                $"Roll: {ToDegrees(eulerArray[0])}, " +
                $"Pitch: {ToDegrees(eulerArray[1])}, " +
                $"Yaw: {ToDegrees(eulerArray[2])}");

            // Update the view center ray.
            (this.ViewFrontRay, this.ViewRightRay) = SphereTracking.EulerAnglesToVectors(eulerArray[2], eulerArray[1], eulerArray[0]);
        }

        public void UpdateScreenCenterRay(Vector3 newRay)
        {
            ScreenCenterRay = newRay;
        }

        private float ToDegrees(float radians)
        {
            return radians * 180.0f / (float)Math.PI;
        }

        // Set the boundaries for position clamping
        public void SetBoundaries(PointF min, PointF max)
        {
            minBoundary = min;
            maxBoundary = max;
            Console.WriteLine($"Head tracking boundaries set: Min({minBoundary.X}, {minBoundary.Y}), Max({maxBoundary.X}, {maxBoundary.Y})");
        }

        // Calculate head position and handle clamping
        public PointF UpdateHeadPosition()
        {
            // Calculate the raw position first
            var newPosition = CalculateHeadPosition();
            
            Console.WriteLine($"Raw calculated position: {newPosition.X}, {newPosition.Y}");
            
            // Check if calculation failed
            if (float.IsNaN(newPosition.X) || float.IsNaN(newPosition.Y))
            {
                Console.WriteLine("Position calculation failed");
                return currentPosition; // Return the last valid position
            }
            
            // Clamp to boundaries
            PointF clampedPosition = new PointF(
                Math.Clamp(newPosition.X, minBoundary.X, maxBoundary.X),
                Math.Clamp(newPosition.Y, minBoundary.Y, maxBoundary.Y)
            );
            
            // If position was clamped, update the screen center ray
            if (clampedPosition != newPosition)
            {
                Console.WriteLine("Position clamped to boundaries");
                
                // Calculate the adjusted ray using the clamped position
                Vector3 adjustedRay = SphereTracking.LocalCoordinatesToRayVector(
                    clampedPosition, ViewFrontRay, ViewRightRay, Radius);
                
                // Update the screen center ray
                ScreenCenterRay = adjustedRay;
                
                Console.WriteLine($"Adjusted screenCenterRay: {ScreenCenterRay.X}, {ScreenCenterRay.Y}, {ScreenCenterRay.Z}");
            }
            
            // Save and return the new position
            currentPosition = clampedPosition;
            return currentPosition;
        }

        // The raw position calculation without clamping
        private PointF CalculateHeadPosition()
        {
            // Calculate the intersection of the viewRay and the sphere
            var screenCenterPoint = SphereTracking.RayPlaneIntersection(
                ScreenCenterRay, ViewFrontRay, Radius);
                
            return SphereTracking.ComputePlaneLocalCoordinates(
                screenCenterPoint, ViewFrontRay, ViewRightRay, Radius);
        }

        // Get the current head position
        public PointF GetCurrentPosition()
        {
            return currentPosition;
        }
    }
}
