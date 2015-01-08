using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace chrisbjohnson.TaskListPlus
{
    /// <summary>
    /// Class containing all extension methods
    /// </summary>
    public static class Extensions
    {
        /// <summary>
        /// Extension method to reset a timer
        /// </summary>
        /// <param name="timer">Timer object to reset</param>
        public static void Reset(this System.Timers.Timer timer)
        {
            // Stopping then starting a timer resets it's elapsed timeout
            timer.Stop();
            timer.Start();
        }
    }
}
