using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Windows.Input;
using System.IO;
using System.Reflection;
using System.Linq.Expressions;
using chrisbjohnson.TaskListPlus.MVVMLib;
using EnvDTE;

namespace chrisbjohnson.TaskListPlus
{
    public class TaskListItemCollectionVM : ObservableObject
    {
        #region Variables

        private EnvDTE80.DTE2 applicationObject;
        private String token;
        private int scope;

        private ObservableCollection<TaskListItemVM> _taskCollection;
        private ObservableCollection<TaskListItemVM> _filteredTaskCollection;

        /// <summary>
        /// Collection of all TaskListItemViewModel
        /// </summary>
        public ObservableCollection<TaskListItemVM> TaskCollection
        {
            get { return _taskCollection; }
            set
            {
                _taskCollection = value;
                OnPropertyChanged("TaskCollection");
            }
        }

        /// <summary>
        /// Filtered collection of TaskListItemViewModel for use with the listview in the GUI
        /// </summary>
        public ObservableCollection<TaskListItemVM> FilteredTaskCollection
        {
            get { return _filteredTaskCollection; }
            set
            {
                _filteredTaskCollection = value;
                OnPropertyChanged("FilteredTaskCollection");
            }
        }
        #endregion

        /// <summary>
        /// Constructor, create a TaskListItemViewModel for each TaskItem in TaskList
        /// </summary>
        /// <param name="ApplicationObject">Reference to the main application</param>
        /// <param name="selectedToken">Selected token from the GUI</param>
        /// <param name="selectedScopeIndex">Selected scope index from the GUI</param>
        public TaskListItemCollectionVM(ref EnvDTE80.DTE2 ApplicationObject, String selectedToken, int selectedScopeIndex)
        {
            // Store the application object as a local reference
            this.applicationObject = ApplicationObject;

            // Initialize the token and scope
            token = "ALL";
            scope = 0;

            GetTasksFromApplicationObject();
            CreateFilteredTaskCollection(selectedToken, selectedScopeIndex);
        }

        /// <summary>
        /// Reads all tasks from the application object passed in
        /// </summary>
        public void GetTasksFromApplicationObject()
        {
            // Instantiate the collection first
            TaskCollection = new ObservableCollection<TaskListItemVM>();

            // Create the full list of task items
            foreach (TaskItem taskItem in applicationObject.ToolWindows.TaskList.TaskItems)
            {
                try
                {
                    // Only show comment task items
                    if (taskItem.Category.Equals("Comment", StringComparison.CurrentCultureIgnoreCase))
                    {
                        TaskCollection.Add(new TaskListItemVM(taskItem));
                    }
                }
                catch
                {
                    // Don't add this item, do nothing
                }
            }
        }

        /// <summary>
        /// Remove the tasks whose filename matches the currently open file
        /// </summary>
        private void RemoveTasksFromCurrentFile()
        {
            for (int i = TaskCollection.Count - 1; i >= 0; i--)
            {
                if (TaskCollection[i].FullFilename.Equals(GetCurrentFile(), StringComparison.CurrentCultureIgnoreCase))
                    _taskCollection.RemoveAt(i);
            }
        }

        /// <summary>
        /// Filter the taskitems based on the token string and/or the scope
        /// </summary>
        /// <param name="token">Token to filter on (e.g. TODO or HACK)</param>
        /// <param name="scope">Scope to filter tasks on (e.g. Project or Class)</param>
        public void CreateFilteredTaskCollection(String token, int scope)
        {
            // Copy passed variables into locals (stored for filtering again later)
            this.token = token;
            this.scope = scope;

            FilteredTaskCollection = FilterTasks();
        }

        /// <summary>
        /// Clears the task collection when changing projects
        /// </summary>
        public void ClearTaskCollections()
        {
            TaskCollection.Clear();
            FilteredTaskCollection.Clear();
        }

        /// <summary>
        /// Filter all tasks - filters based on the locally stored variables (can be re-run internally)
        /// </summary>
        /// <returns>Collection of all task list items</returns>
        private ObservableCollection<TaskListItemVM> FilterTasks()
        {
            List<String> projectFiles = new List<String>();
            String currentFile = "";

            // Create a temporary list (can't change observablecollection from different thread it was created on, we call this from a timer)
            ObservableCollection<TaskListItemVM> tempFilteredTaskCollection = new ObservableCollection<TaskListItemVM>();

            // Before we start looping through the tasks check the scope
            // If it's Project or Class, get a list of filenames of the files that match
            if (scope == (int)TaskListControlVM.taskScope.Project)
            {
                projectFiles = GetFilesInProject();
            }
            else if (scope == (int)TaskListControlVM.taskScope.Class)
            {
                currentFile = GetCurrentFile();
            }

            // Add the item if it has the matching token
            foreach (TaskListItemVM taskItem in TaskCollection)
            {
                // First check the token to look for a match ("ALL" is a special case, with an obvious use)
                if (token.Equals("ALL") || taskItem.Token.Equals(token, StringComparison.CurrentCultureIgnoreCase))
                {
                    // Now check the scope option
                    // Solution will just add the item - it doesn't filter at all
                    if (scope == (int)TaskListControlVM.taskScope.Solution)
                    {
                        tempFilteredTaskCollection.Add(taskItem);
                    }
                    // Project needs to check the taskItem file versus files in the project for matches
                    else if (scope == (int)TaskListControlVM.taskScope.Project)
                    {
                        // If the list of project files contains the taskitems filename then add it to the filtered list
                        if (projectFiles.Contains(taskItem.FullFilename))
                            tempFilteredTaskCollection.Add(taskItem);
                    }
                    // Class looks for task items matching only the class open in the editor
                    else if (scope == (int)TaskListControlVM.taskScope.Class)
                    {
                        if (currentFile.Equals(taskItem.FullFilename, StringComparison.CurrentCultureIgnoreCase))
                            tempFilteredTaskCollection.Add(taskItem);
                    }
                }
            }

            return tempFilteredTaskCollection;
        }

        /// <summary>
        /// Returns the current filename of the file open in the editor
        /// </summary>
        /// <returns>current filename of the file open in the editor</returns>
        public string GetCurrentFile()
        {
            try
            {
                return applicationObject.ActiveDocument.FullName;
            }
            catch
            {
                // If this file doens't exist, return a blank string
                return "";
            }

        }

        /// <summary>
        /// Gets a list of all files contained in the current open project
        /// </summary>
        /// <returns>List of all files in the current project</returns>
        private List<string> GetFilesInProject()
        {
            List<String> files = new List<String>();

            // Get the active project
            Project activeProject = GetActiveProject();

            try
            {

                if (activeProject != null)
                {
                    // Now loop through all the project items and add the filenames to the projectfiles list
                    foreach (ProjectItem item in activeProject.ProjectItems)
                    {
                        if (item.FileNames[0].EndsWith("\\"))
                        {
                            // This is a folder, so add all the items in the subfolder (doesn't really matter if some of them are not really appropriate
                            files.AddRange(Directory.GetFiles(item.FileNames[0], "*", SearchOption.AllDirectories));
                        }
                        else
                        {
                            // For some reason, the filename is a list, with the actual filename at 0 index
                            files.Add(item.FileNames[0]);
                        }
                    }
                }
            }
            catch
            {
                // I don't think it matters if we ignore this error
            }

            return files;
        }

        /// <summary>
        /// Returns the active project as a project object
        /// </summary>
        /// <returns>Project object</returns>
        internal Project GetActiveProject()
        {
            Project activeProject = null;

            Array activeSolutionProjects = applicationObject.ActiveSolutionProjects as Array;
            if (activeSolutionProjects != null && activeSolutionProjects.Length > 0)
            {
                activeProject = activeSolutionProjects.GetValue(0) as Project;
            }

            return activeProject;
        }
    }
}
