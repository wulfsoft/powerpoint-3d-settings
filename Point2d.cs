namespace Ppt3dSettingsFinder
{
    /// <summary>
    /// Represents a point in a two-dimensional coordinate system.
    /// </summary>
    public class Point2d
    {
        public double X { get; set; }
        public double Y { get; set; }

        public Point2d(double x, double y)
        {
            this.X = x;
            this.Y = y;
        }

        public override string ToString()
        {
            return string.Format("({0}, {1})", X, Y);
        }
    }
}