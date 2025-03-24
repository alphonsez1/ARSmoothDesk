using System;
using System.Numerics;

public static class SphereTracking
{
    public static (Vector3 frontVector, Vector3 rightVector) EulerAnglesToVectors(float yaw, float pitch, float roll)
    {
        // Convert to radians once
        float yawRad = ToRadians(yaw);
        float pitchRad = ToRadians(pitch);
        float rollRad = ToRadians(roll);
        
        // Calculate sines and cosines once
        float sinYaw = (float)Math.Sin(yawRad);
        float cosYaw = (float)Math.Cos(yawRad);
        float sinPitch = (float)Math.Sin(pitchRad);
        float cosPitch = (float)Math.Cos(pitchRad);
        float sinRoll = (float)Math.Sin(rollRad);
        float cosRoll = (float)Math.Cos(rollRad);
        
        // Calculate front vector
        Vector3 frontVector = new Vector3(
            -sinYaw * cosPitch,
            -sinPitch,
            cosYaw * cosPitch
        );
        
        // Calculate right vector
        Vector3 rightVector = new Vector3(
            cosYaw * cosRoll + sinYaw * sinPitch * sinRoll,
            cosPitch * sinRoll,
            sinYaw * cosRoll - cosYaw * sinPitch * sinRoll
        );
        
        return (frontVector, rightVector);
    }
    
    // 2. Find the intersection point where a ray hits a tangent plane
    public static Vector3 RayPlaneIntersection(Vector3 originalRay, Vector3 planeNormal, float sphereRadius)
    {
        // Compute the denominator (dot product of ray and plane normal)
        float denominator = Vector3.Dot(originalRay, planeNormal);
        
        // Check if the ray is parallel to the plane
        if (Math.Abs(denominator) < 1e-6f)
        {
            // The ray is parallel to the plane, no intersection
            return new Vector3(float.NaN, float.NaN, float.NaN);
        }
        
        // Calculate the parameter t for the ray equation
        float t = sphereRadius / denominator;
        
        // Calculate the intersection point
        return originalRay * t;
    }

    private static (Vector3 e1, Vector3 e2) CreateTangentPlaneCoordinateSystem(Vector3 planeNormal, Vector3 rightDirection)
    {
        Vector3 e1 = rightDirection;

        // Normalize e1
        e1 = Vector3.Normalize(e1);
        
        // Create e2 by rotating e1 by 90 degrees counter-clockwise on the plane
        // This is done by taking the cross product of the normal and e1
        Vector3 e2 = Vector3.Cross(planeNormal, e1);
        
        // Normalize e2 to ensure it's a unit vector
        e2 = Vector3.Normalize(e2);

        Console.WriteLine($"e1: {e1}, e2: {e2}");
        
        return (e1, e2);
    }

    // 3. Compute local coordinates on the tangent plane
    public static PointF ComputePlaneLocalCoordinates(Vector3 point, Vector3 planeNormal, Vector3 rightDirection, float sphereRadius)
    {
        // The tangent point on the sphere
        Vector3 tangentPoint = planeNormal * sphereRadius;
        
        // Create coordinate system with roll
        var (e1, e2) = CreateTangentPlaneCoordinateSystem(planeNormal, rightDirection);
        
        // Vector from tangent point to intersection point
        Vector3 w = point - tangentPoint;
        
        // Project onto the basis vectors to get local coordinates
        float u = Vector3.Dot(w, e1);
        float v = Vector3.Dot(w, e2);
        
        return new PointF(u, v);
    }

    // 4. Convert local coordinates back to a ray vector
    public static Vector3 LocalCoordinatesToRayVector(PointF localCoords, Vector3 planeNormal, Vector3 rightDirection, float sphereRadius)
    {
        // The tangent point on the sphere
        Vector3 tangentPoint = planeNormal * sphereRadius;
        
        // Create coordinate system with roll
        var (e1, e2) = CreateTangentPlaneCoordinateSystem(planeNormal, rightDirection);
        
        // Compute the 3D point from local coordinates
        Vector3 point = tangentPoint + e1 * localCoords.X + e2 * localCoords.Y;
        
        // Return the normalized ray vector
        return Vector3.Normalize(point);
    }
    
    private static float ToRadians(float degrees)
    {
        return degrees * (float)Math.PI / 180.0f;
    }

    public static float CalculateDistanceToDisplay(float screenWidth, float screenHeight, float diagonalFovDegrees)
    {
        // Convert diagonal FOV from degrees to radians
        float diagonalFovRadians = diagonalFovDegrees * (MathF.PI / 180f);
        
        // Calculate screen diagonal in pixels
        float screenDiagonal = MathF.Sqrt(screenWidth * screenWidth + screenHeight * screenHeight);
        
        // Calculate distance using the tangent relationship
        // distance = (diagonal / 2) / tan(FOV / 2)
        float distance = (screenDiagonal / 2) / MathF.Tan(diagonalFovRadians / 2);
        
        return distance;
    }
}