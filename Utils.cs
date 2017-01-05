using System;

namespace Ppt3dSettingsFinder
{
    /// <summary>
    /// Provides utility methods for the application.
    /// </summary>
    public static class Utils
    {
        /// <summary>
        /// Converts a number from degrees to radians.
        /// </summary>
        /// <param name="value">An angle measured in degrees.</param>
        public static double DegreesToRadians(double value)
        {
            return value * Math.PI / 180;
        }
    }
}