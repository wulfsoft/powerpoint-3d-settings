using System;

namespace Ppt3dSettingsFinder
{
    /// <summary>
    /// Represents a 4x4 affine transformation matrix used for transformation in 3D space.
    /// </summary>
    public class TransformationMatrix
    {
        /// <summary>
        /// Gets or sets the value of the first row and first column of the matrix.
        /// </summary>
        public double M11 { get; set; }

        /// <summary>
        /// Gets or sets the value of the first row and second column of the matrix.
        /// </summary>
        public double M12 { get; set; }

        /// <summary>
        /// Gets or sets the value of the first row and third column of the matrix.
        /// </summary>
        public double M13 { get; set; }

        /// <summary>
        /// Gets or sets the value of the first row and fourth column of the matrix.
        /// </summary>
        public double M14 { get; set; }

        /// <summary>
        /// Gets or sets the value of the second row and first column of the matrix.
        /// </summary>
        public double M21 { get; set; }

        /// <summary>
        /// Gets or sets the value of the second row and second column of the matrix.
        /// </summary>
        public double M22 { get; set; }

        /// <summary>
        /// Gets or sets the value of the second row and third column of the matrix.
        /// </summary>
        public double M23 { get; set; }

        /// <summary>
        /// Gets or sets the value of the second row and fourth column of the matrix.
        /// </summary>
        public double M24 { get; set; }

        /// <summary>
        /// Gets or sets the value of the third row and first column of the matrix.
        /// </summary>
        public double M31 { get; set; }

        /// <summary>
        /// Gets or sets the value of the third row and second column of the matrix.
        /// </summary>
        public double M32 { get; set; }

        /// <summary>
        /// Gets or sets the value of the third row and third column of the matrix.
        /// </summary>
        public double M33 { get; set; }

        /// <summary>
        /// Gets or sets the value of the third row and fourth column of the matrix.
        /// </summary>
        public double M34 { get; set; }

        public TransformationMatrix(
            double m11 = 1, double m12 = 0, double m13 = 0, double m14 = 0,
            double m21 = 0, double m22 = 1, double m23 = 0, double m24 = 0,
            double m31 = 0, double m32 = 0, double m33 = 1, double m34 = 0)
        {
            M11 = m11;
            M12 = m12;
            M13 = m13;
            M14 = m14;
            M21 = m21;
            M22 = m22;
            M23 = m23;
            M24 = m24;
            M31 = m31;
            M32 = m32;
            M33 = m33;
            M34 = m34;
        }

        public override string ToString()
        {
            return string.Format("[[{0}, {1}, {2}, {3}], [{4}, {5}, {6}, {7}], [{8}, {9}, {10}, {11}]]",
                M11, M12, M13, M14, M21, M22, M23, M24, M31, M32, M33, M34);
        }

        /// <summary>
        /// Appends the specified matrix to this matrix.
        /// </summary>
        public void Multiply(TransformationMatrix transform)
        {
            var m11 = M11 * transform.M11 + M12 * transform.M21 + M13 * transform.M31;
            var m12 = M11 * transform.M12 + M12 * transform.M22 + M13 * transform.M32;
            var m13 = M11 * transform.M13 + M12 * transform.M23 + M13 * transform.M33;
            var m14 = M11 * transform.M14 + M12 * transform.M24 + M13 * transform.M34 + M14;
            var m21 = M21 * transform.M11 + M22 * transform.M21 + M23 * transform.M31;
            var m22 = M21 * transform.M12 + M22 * transform.M22 + M23 * transform.M32;
            var m23 = M21 * transform.M13 + M22 * transform.M23 + M23 * transform.M33;
            var m24 = M21 * transform.M14 + M22 * transform.M24 + M23 * transform.M34 + M24;
            var m31 = M31 * transform.M11 + M32 * transform.M21 + M33 * transform.M31;
            var m32 = M31 * transform.M12 + M32 * transform.M22 + M33 * transform.M32;
            var m33 = M31 * transform.M13 + M32 * transform.M23 + M33 * transform.M33;
            var m34 = M31 * transform.M14 + M32 * transform.M24 + M33 * transform.M34 + M34;

            M11 = m11;
            M12 = m12;
            M13 = m13;
            M14 = m14;
            M21 = m21;
            M22 = m22;
            M23 = m23;
            M24 = m24;
            M31 = m31;
            M32 = m32;
            M33 = m33;
            M34 = m34;
        }

        /// <summary>
        /// Applies a rotation of the specified angle about the x axis.
        /// </summary>
        /// <param name="angleInRadians">The rotation angle in radians.</param>
        public void RotateX(double angleInRadians)
        {
            var rotationMatrix = new TransformationMatrix(
                1, 0, 0, 0,
                0, Math.Cos(angleInRadians), -Math.Sin(angleInRadians), 0,
                0, Math.Sin(angleInRadians), Math.Cos(angleInRadians), 0);

            Multiply(rotationMatrix);
        }

        /// <summary>
        /// Applies a rotation of the specified angle about the y axis.
        /// </summary>
        /// <param name="angleInRadians">The rotation angle in radians.</param>
        public void RotateY(double angleInRadians)
        {
            var rotationMatrix = new TransformationMatrix(
                Math.Cos(angleInRadians), 0, Math.Sin(angleInRadians), 0,
                0, 1, 0, 0,
                -Math.Sin(angleInRadians), 0, Math.Cos(angleInRadians), 0);

            Multiply(rotationMatrix);
        }

        /// <summary>
        /// Applies a rotation of the specified angle about the z axis.
        /// </summary>
        /// <param name="angleInRadians">The rotation angle in radians.</param>
        public void RotateZ(double angleInRadians)
        {
            var rotationMatrix = new TransformationMatrix(
                Math.Cos(angleInRadians), -Math.Sin(angleInRadians), 0, 0,
                Math.Sin(angleInRadians), Math.Cos(angleInRadians), 0, 0,
                0, 0, 1, 0);

            Multiply(rotationMatrix);
        }

        /// <summary>
        /// Transforms the specified point by this matrix.
        /// </summary>
        public Point3d Transform(Point3d point)
        {
            return new Point3d(
                M11 * point.X + M12 * point.Y + M13 * point.Z + M14,
                M21 * point.X + M22 * point.Y + M23 * point.Z + M24,
                M31 * point.X + M32 * point.Y + M33 * point.Z + M34);
        }

        /// <summary>
        /// Projects a point from 3D to 2D space.
        /// </summary>
        /// <param name="point">The 3D point to project to 2D.</param>
        /// <param name="angleOfViewInRadians">The camera vertical angle of view (in radians).</param>
        public Point2d Project(Point3d point, double angleOfViewInRadians)
        {
            var s = 1 / Math.Tan(angleOfViewInRadians * 0.5);
            if (double.IsInfinity(s))
                s = 1;
            var x = s * point.X;
            var y = s * point.Y;
            var z = point.Z;

            // Camera transform
            if (angleOfViewInRadians == 0)
                z = 1;
            else
                z -= s * 2 * Math.PI;
            
            if (z != 0)
            {
                x = x / Math.Abs(z) * 2 * Math.PI;
                y = y / Math.Abs(z) * 2 * Math.PI;
            }
            
            return new Point2d(x, y);
        }
    }
}
