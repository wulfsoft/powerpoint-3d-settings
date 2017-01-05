using System;

namespace Ppt3dSettingsFinder
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // Specify the four vertices of the target quadrilateral on the PowerPoint slide
            var topLeft = new Point2d(1.94, 1.82);
            var topRight = new Point2d(6.62, 0.4);
            var bottomRight = new Point2d(7.54, 3.85);
            var bottomLeft = new Point2d(2.95, 5.45);

            // Specify the size of the rectangle to transform into the quadrilateral defined above
            var rectangleWidth = 1440;
            var rectangleHeight = 900;

            Console.WriteLine(string.Format(
                "Searching for the optimal 3D settings for transforming a rectangle into a quadrilateral defined by the following 4 points: {0}, {1}, {2}, {3}",
                topLeft, topRight, bottomRight, bottomLeft));

            // Start the search for the optimal settings
            var task = Ppt3dSettingsFinder.FindOptimalShapeSettings(topLeft, topRight, bottomRight, bottomLeft, rectangleWidth, rectangleHeight);
            task.Wait();
            Console.WriteLine(string.Format(
                "Found shape settings:\r\nWidth: {0:0.##}\r\nHeight: {1:0.##}\r\nX Rotation: {2:0.#}\r\nY Rotation: {3:0.#}\r\nZ Rotation: {4:0.#}\r\nPerspective: {5:0.#}\r\n(Estimate: {6:0.######})",
                task.Result.Width, task.Result.Height, task.Result.XRotation, task.Result.YRotation, task.Result.ZRotation,
                task.Result.Perspective, task.Result.Estimate));
        }
    }
}