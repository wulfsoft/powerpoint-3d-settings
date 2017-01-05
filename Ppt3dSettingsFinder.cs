using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Ppt3dSettingsFinder
{
    public class Ppt3dSettingsFinder
    {
        /// <summary>
        /// Searches for the optimal shape settings that transform a rectangle of size <paramref name="rectangleWidth"/> by
        /// <paramref name="rectangleWidth"/> into a quadrilateral defined by the four vertices <paramref name="topLeft"/>,
        /// <paramref name="topRight"/>, <paramref name="bottomRight"/> and <paramref name="bottomLeft"/>.
        /// </summary>
        public static async Task<ShapeSettings> FindOptimalShapeSettings(
            Point2d topLeft, Point2d topRight, Point2d bottomRight, Point2d bottomLeft,
            double rectangleWidth, double rectangleHeight)
        {
            ShapeSettings shapeSettings = null;

            // Find a reasonable interval of width values
            var vertices = new Point2d[] { topLeft, topRight, bottomRight, bottomLeft };
            var width = vertices.Max(p => p.X) - vertices.Min(p => p.X);
            var height = vertices.Max(p => p.Y) - vertices.Min(p => p.Y);
            var aspectRatio = rectangleWidth / rectangleHeight;
            if (width / aspectRatio < height)
                width = height * aspectRatio;
            double minWidth = Math.Round(width * 0.3, 2);
            double maxWidth = Math.Round(width * 1.1, 2);
            double widthStep = Math.Round((maxWidth - minWidth) / 50, 2);

            // Check "X Rotation" values from 0 to 360 degrees
            double minXRotation = 0;
            double maxXRotation = 360;

            // Check "Y Rotation" values from 0 to 360 degrees
            double minYRotation = 0;
            double maxYRotation = 360;

            // Check "Z Rotation" values from 0 to 360 degrees
            double minZRotation = 0;
            double maxZRotation = 360;

            // Check "Perspective" values from 0 to 360 degrees
            double minPerspective = 0;
            double maxPerspective = 120;

            // If you don't get a satisfying solution, try increasing or decreasing the step value
            double rotationStep = 30;

            // Since it's not feasible to check all combinations of possible width, xRotation, yRotation, zRotation
            // and perspective values, a nested interval approach is used to get close to the optimal solution
            do
            {
                shapeSettings = await FindOptimalShapeSettings(
                    topLeft, topRight, bottomRight, bottomLeft,
                    rectangleWidth, rectangleHeight,
                    minWidth, maxWidth, widthStep,
                    minXRotation, maxXRotation, rotationStep,
                    minYRotation, maxYRotation, rotationStep,
                    minZRotation, maxZRotation, rotationStep,
                    minPerspective, maxPerspective, rotationStep);                

                if (rotationStep / 2 >= 0.1) {
                    // Reduce the size of the intervals and steps with each iteration

                    rotationStep = Math.Ceiling(rotationStep / 2 * 10) / 10;
                    widthStep = Math.Ceiling(widthStep / 2 * 100) / 100;

                    minWidth = Math.Max(shapeSettings.Width - 8 * widthStep, 0);
                    maxWidth = shapeSettings.Width + 8 * widthStep;

                    minXRotation = shapeSettings.XRotation - 4 * rotationStep;
                    maxXRotation = shapeSettings.XRotation + 4 * rotationStep;
                    
                    minYRotation = shapeSettings.YRotation - 4 * rotationStep;
                    maxYRotation = shapeSettings.YRotation + 4 * rotationStep;
                    
                    minZRotation = shapeSettings.ZRotation - 4 * rotationStep;
                    maxZRotation = shapeSettings.ZRotation + 4 * rotationStep;

                    minPerspective = Math.Max(shapeSettings.Perspective - 6 * rotationStep, 0);
                    maxPerspective = Math.Min(shapeSettings.Perspective + 6 * rotationStep, 120);
                } else {
                    // Optimal solution found
                    break;
                }
            } while (true);

            return shapeSettings;
        }

        private static async Task<ShapeSettings> FindOptimalShapeSettings(
            Point2d topLeft, Point2d topRight, Point2d bottomRight, Point2d bottomLeft,
            double rectangleWidth, double rectangleHeight,
            double minWidth, double maxWidth, double widthStep,
            double minXRotation, double maxXRotation, double xRotationStep,
            double minYRotation, double maxYRotation, double yRotationStep,
            double minZRotation, double maxZRotation, double zRotationStep,
            double minPerspective, double maxPerspective, double perspectiveStep)
        {
            var shapeSettings = new ShapeSettings();
            var aspectRatio = rectangleWidth / rectangleHeight;                
            var lockToken = new Object();

            // Check all combinations of allowed width, xRotation, yRotation, zRotation and perspective values
            var tasks = new List<Task>();
            for (var width = minWidth; width < maxWidth; width += widthStep)
            {
                // Parallelize the search into tasks
                tasks.Add(new Task(currentWidth => {
                    for (double xRotation = minXRotation; xRotation < maxXRotation; xRotation += xRotationStep)
                    {
                        for (double yRotation = minYRotation; yRotation < maxYRotation; yRotation += yRotationStep)
                        {
                            for (double zRotation = minZRotation; zRotation < maxZRotation; zRotation += zRotationStep)
                            {
                                for (double perspective = minPerspective; perspective < maxPerspective; perspective += perspectiveStep)
                                {
                                    // Check how close the current combination of input values gets to the intended target
                                    var estimate = Evaluate(
                                        topLeft, topRight, bottomRight, bottomLeft,
                                        (double)currentWidth, (double)currentWidth / aspectRatio,
                                        xRotation, yRotation, zRotation, perspective);

                                    lock (lockToken)
                                    {
                                        if (estimate < shapeSettings.Estimate)
                                        {
                                            // A better solution was found

                                            shapeSettings.Width = (double)currentWidth;
                                            shapeSettings.Height = (double)currentWidth / aspectRatio;
                                            shapeSettings.XRotation = xRotation;
                                            shapeSettings.YRotation = yRotation;
                                            shapeSettings.ZRotation = zRotation;
                                            shapeSettings.Perspective = perspective;
                                            shapeSettings.Estimate = estimate;

                                            if (shapeSettings.XRotation < 0)
                                                shapeSettings.XRotation += 360;
                                            else if (shapeSettings.XRotation > 360)
                                                shapeSettings.XRotation -= 360;
                                            if (shapeSettings.YRotation < 0)
                                                shapeSettings.YRotation += 360;
                                            else if (shapeSettings.YRotation > 360)
                                                shapeSettings.YRotation -= 360;
                                            if (shapeSettings.ZRotation < 0)
                                                shapeSettings.ZRotation += 360;
                                            else if (shapeSettings.ZRotation > 360)
                                                shapeSettings.ZRotation -= 360;                       
                                        }
                                    }
                                }
                            }
                        }
                    }
                }, width));
            }

            // Start the tasks and wait for them to complete
            tasks.ForEach(t => t.Start());
            await Task.WhenAll(tasks.ToArray());

            return shapeSettings;
        }

        private static double Evaluate(
            Point2d topLeft, Point2d topRight, Point2d bottomRight, Point2d bottomLeft,
            double rectangleWidth, double rectangleHeight, double xRotation, double yRotation, double zRotation, double perspective)
        {
            var transform = new TransformationMatrix();
            transform.RotateZ(Utils.DegreesToRadians(-zRotation));  // Called "Z Rotation" in PowerPoint
            transform.RotateX(Utils.DegreesToRadians(-yRotation));  // Called "Y Rotation" in PowerPoint
            transform.RotateY(Utils.DegreesToRadians(-xRotation));  // Called "X Rotation" in PowerPoint

            var perspectiveInRadians = Utils.DegreesToRadians(perspective);
            var topLeftProjected = transform.Project(transform.Transform(new Point3d(-rectangleWidth / 2, -rectangleHeight / 2, 0)), perspectiveInRadians);
            var topRightProjected = transform.Project(transform.Transform(new Point3d(rectangleWidth / 2, -rectangleHeight / 2, 0)), perspectiveInRadians);
            var bottomRightProjected = transform.Project(transform.Transform(new Point3d(rectangleWidth / 2, rectangleHeight / 2, 0)), perspectiveInRadians);
            var bottomLeftProjected = transform.Project(transform.Transform(new Point3d(-rectangleWidth / 2, rectangleHeight / 2, 0)), perspectiveInRadians);
            
            var projectedPoints = new Point2d[] { topLeftProjected, topRightProjected, bottomRightProjected, bottomLeftProjected };
            var minXProjected = projectedPoints.Min(p => p.X);
            var minYProjected = projectedPoints.Min(p => p.Y);

            var points = new Point2d[] { topLeft, topRight, bottomRight, bottomLeft };
            var minX = points.Min(p => p.X);
            var minY = points.Min(p => p.Y);

            // Return the sum of squared distances between the transformed points and the target points
            return
                Math.Pow((topLeftProjected.X - minXProjected) - (topLeft.X - minX), 2) +
                Math.Pow((topLeftProjected.Y - minYProjected) - (topLeft.Y - minY), 2) +
                Math.Pow((topRightProjected.X - minXProjected) - (topRight.X - minX), 2) +
                Math.Pow((topRightProjected.Y - minYProjected) - (topRight.Y - minY), 2) +
                Math.Pow((bottomRightProjected.X - minXProjected) - (bottomRight.X - minX), 2) +
                Math.Pow((bottomRightProjected.Y - minYProjected) - (bottomRight.Y - minY), 2) +
                Math.Pow((bottomLeftProjected.X - minXProjected) - (bottomLeft.X - minX), 2) +
                Math.Pow((bottomLeftProjected.Y - minYProjected) - (bottomLeft.Y - minY), 2);
        }
    }
}