namespace Ppt3dSettingsFinder
{
    /// <summary>
    /// Size and 3D settings for a shape on a PowerPoint slide.
    /// </summary>
    public class ShapeSettings {
        /// <summary>
        /// The "Width" value of the shape.
        /// </summary>
        public double Width { get; set; }

        /// <summary>
        /// The "Height" value of the shape.
        /// </summary>
        public double Height { get; set; }

        /// <summary>
        /// The "X Rotation" value of the shape.
        /// </summary>
        public double XRotation { get; set; }
        
        /// <summary>
        /// The "Y Rotation" value of the shape.
        /// </summary>
        public double YRotation { get; set; }

        /// <summary>
        /// The "Z Rotation" value of the shape.
        /// </summary>
        public double ZRotation { get; set; }

        /// <summary>
        /// The "Perspective" value of the shape.
        /// </summary>
        public double Perspective { get; set; }

        /// <summary>
        /// A value representing how close a 3D transformation of a rectangle with this settings
        /// gets to the target quadrilateral; the closer this value is to 0, the more accurate is the
        /// result of the transformation.
        /// </summary>
        public double Estimate { get; set; }

        public ShapeSettings()
        {
            Estimate = double.MaxValue;
        }
    }
}