using System;
using System.Text;

namespace chrisbjohnson.TaskListPlus
{
    /// <summary>
    /// TaskListItem holds all data for each Tas
    /// </summary>
    public class TaskListItem
    {
        private String _token;
        private String _description;
        private String _filename;
        private String _fullFilename;        
        private int _line;

        /// <summary>
        /// Token for the task item e.g. TODO
        /// </summary>
        public String Token
        {
            get { return _token; }
            set { _token = value; }
        }
        
        /// <summary>
        /// Description of the task
        /// </summary>
        public String Description
        {
            get { return _description; }
            set { _description = value; }
        }

        /// <summary>
        /// Filename containing the task
        /// </summary>
        public String Filename
        {
            get { return _filename; }
            set { _filename = value; }
        }
        
        /// <summary>
        /// Full filename including directory path
        /// </summary>
        public String FullFilename
        {
            get { return _fullFilename; }
            set { _fullFilename = value; }
        }

        /// <summary>
        /// Line to locate the task
        /// </summary>
        public int Line
        {
            get { return _line; }
            set { _line = value; }
        }
    }
}
