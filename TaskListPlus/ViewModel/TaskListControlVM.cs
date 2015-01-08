using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Input;
using chrisbjohnson.TaskListPlus.MVVMLib;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Shell;
using System.Timers;
using EnvDTE;
using EnvDTE80;

namespace chrisbjohnson.TaskListPlus
{
    public class TaskListControlVM : ObservableObject
    {
        private String openFile;
        private TaskListItemCollectionVM _taskViewModel;
        private TokenVM _tokenViewModel;
        private int _selectedScopeIndex;
        private String _selectedToken;

        private EnvDTE.SolutionEvents solutionEvents;
        private EnvDTE.TextEditorEvents textEditorEvents;
        private EnvDTE.WindowEvents windowEvents;

        /// <summary>
        /// Delay before firing the command to re-read tasks
        /// Stops the tasks being read every time something is typed
        /// Only when no text has changed in the editor for the UPDATE_DELAY duration will the updates occur.
        /// Also, the application's task list doesn't update for at least 1 second after changes are made.
        /// </summary>
        private System.Timers.Timer taskUpdateDelayTimer;
        private const int UPDATE_DELAY = 2000;

        /// <summary>
        /// Delay before reading tasks after a new solution opens
        /// There is an in-built delay in the standard task list before it gets updated
        /// </summary>
        private System.Timers.Timer solutionOpenedDelayTimer;
        private const int SOLUTION_OPENED_DELAY = 3000;
        private int retrySolutionOpenedCount;
        private int currentListCount;

        /// <summary>
        /// Enum representing the different versions of scope of tasks being filtered
        /// </summary>
        public enum taskScope
        {
            Solution = 0,
            Project = 1,
            Class = 2
        }

        public static readonly string[] scopeComboChoices = { "Solution", "Project", "Class" };

        /// <summary>
        /// View model for the task item collection
        /// </summary>
        public TaskListItemCollectionVM TaskViewModel
        {
            get { return _taskViewModel; }
            set { _taskViewModel = value; }
        }

        /// <summary>
        /// View model for the tokens available
        /// </summary>
        public TokenVM TokenViewModel
        {
            get { return _tokenViewModel; }
            set { _tokenViewModel = value; }
        }

        /// <summary>
        /// Index of the selected scope combobox on the GUI
        /// </summary>
        public int SelectedScopeIndex
        {
            get { return _selectedScopeIndex; }
            set
            {
                _selectedScopeIndex = value;
                OnPropertyChanged("SelectedScopeIndex");
                RefilterTaskList();
            }
        }

        /// <summary>
        /// String object of the current selected token in the GUI
        /// </summary>
        public String SelectedToken
        {
            get { return _selectedToken; }
            set
            {
                _selectedToken = value;
                OnPropertyChanged("SelectedToken");
                RefilterTaskList();
            }
        }

        /// <summary>
        /// Constructor, creates each sub view model (one for tasks and one for tokens)
        /// Also initializes the two comboboxes to default values
        /// </summary>
        /// <param name="applicationObject">Reference to the main application object</param>
        public TaskListControlVM()
        {
            // Set default values for the two comboboxes
            SelectedScopeIndex = (int)taskScope.Solution;
            SelectedToken = "ALL";

            taskUpdateDelayTimer = new System.Timers.Timer(UPDATE_DELAY);
            taskUpdateDelayTimer.AutoReset = false;
            taskUpdateDelayTimer.Elapsed += new ElapsedEventHandler(taskUpdateTimer_Elapsed);

            solutionOpenedDelayTimer = new System.Timers.Timer(SOLUTION_OPENED_DELAY);
            solutionOpenedDelayTimer.AutoReset = false;
            solutionOpenedDelayTimer.Elapsed += new ElapsedEventHandler(solutionOpenedTimer_Elapsed);

            EnvDTE80.DTE2 applicationObject = Package.GetGlobalService(typeof(SDTE)) as DTE2;

            AddApplicationObjectHandlers(applicationObject);

            // Instantiate each view mode
            TaskViewModel = new TaskListItemCollectionVM(ref applicationObject, SelectedToken, SelectedScopeIndex);
            TokenViewModel = new TokenVM(ref applicationObject);

            // Read the current open file on startup
            openFile = TaskViewModel.GetCurrentFile();
        }

        private EnvDTE80.DTE2 GetApplicationObject()
        {
            throw new NotImplementedException();
        }

        private void AddApplicationObjectHandlers(EnvDTE80.DTE2 _applicationObject)
        {
            // Setup all the solution events
            solutionEvents = (EnvDTE.SolutionEvents)_applicationObject.Events.SolutionEvents;
            solutionEvents.BeforeClosing += new _dispSolutionEvents_BeforeClosingEventHandler(solutionEvents_BeforeClosing);
            solutionEvents.Opened += new _dispSolutionEvents_OpenedEventHandler(solutionEvents_Opened);

            // Setup all text editor events
            textEditorEvents = (EnvDTE.TextEditorEvents)_applicationObject.Events.TextEditorEvents;
            textEditorEvents.LineChanged += new _dispTextEditorEvents_LineChangedEventHandler(textEditorEvents_LineChanged);

            // Setup all window events
            windowEvents = (EnvDTE.WindowEvents)_applicationObject.Events.WindowEvents;
            windowEvents.WindowActivated += new _dispWindowEvents_WindowActivatedEventHandler(windowEvents_WindowActivated);
        }

        /// <summary>
        /// Filters task list based on the new options selected from the GUI
        /// </summary>
        internal void RefilterTaskList()
        {
            if (TaskViewModel == null) return;

            TaskViewModel.CreateFilteredTaskCollection(SelectedToken, SelectedScopeIndex);
        }

        /// <summary>
        /// Filters task list when the current open file changes
        /// </summary>
        internal void RefilterTaskListOnFileChange()
        {
            if (SelectedToken == null) return;

            // Check to see if the file has changed. If it has, update the variable and call method to refilter
            if (!openFile.Equals(TaskViewModel.GetCurrentFile(), StringComparison.CurrentCultureIgnoreCase))
            {
                openFile = TaskViewModel.GetCurrentFile();
                TaskViewModel.CreateFilteredTaskCollection(SelectedToken, SelectedScopeIndex);
            }
        }

        public void ReadAllTasksAndFilter()
        {
            TaskViewModel.GetTasksFromApplicationObject();
            TaskViewModel.CreateFilteredTaskCollection(SelectedToken, SelectedScopeIndex);
        }

        #region Commands

        public void AddToken()
        {
            // Open the token editor window
            TokenEditor tokenEditor = new TokenEditor();
            tokenEditor.ShowDialog();

            // If the OK button was pressed, add the token to the list
            if (tokenEditor.DialogResult == true)
            {
                TokenViewModel.AddToken(tokenEditor.Token);
            }

            // Select the new token from the list
            SelectedToken = tokenEditor.Token;

            // Force update of tokens
            TaskViewModel.GetTasksFromApplicationObject();
            TaskViewModel.CreateFilteredTaskCollection(SelectedToken, SelectedScopeIndex);
        }

        /// <summary>
        /// Command for adding a token button the GUI
        /// Token must have only alphanumeric characters, _, $, ( or )
        /// </summary>
        public void RemoveToken()
        {
            // TODO cannot be removed for some reason
            if (!CanRemoveToken())
            {
                System.Windows.MessageBox.Show(String.Format("Cannot remove {0} token.", SelectedToken), "TaskListPlus");
                return;
            }
            
            // Add the new token to the list and control
            TokenViewModel.RemoveToken(SelectedToken);

            // Select the default "ALL" token from the list
            SelectedToken = "ALL";

            // Force update of tokens
            TaskViewModel.GetTasksFromApplicationObject();
            TaskViewModel.CreateFilteredTaskCollection(SelectedToken, SelectedScopeIndex);
        }

        /// <summary>
        /// Checks to see if token can be deleted
        /// </summary>
        /// <returns>True if token can be removed</returns>
        private bool CanRemoveToken()
        {
            switch (SelectedToken.ToUpper())
            {
                case "ALL":
                    return false;

                case "TODO":
                    return false;

                case "HACK":
                    return false;

                case "UNDONE":
                    return false;

                case "UNRESOLVEDMERGECONFLICT":
                    return false;

                default:
                    return true;
            }
        }

        #endregion

        #region Environment Events

        /// <summary>
        /// Called when a new solution is opened
        /// </summary>
        public void solutionEvents_Opened()
        {
            // Reset the task update timer - it takes a small delay from opening solution to task list being updated
            solutionOpenedDelayTimer.Reset();

            // Reset the retry count - we want to re-read this a number of times to allow for larger projects opening
            retrySolutionOpenedCount = 0;
            currentListCount = 0;
        }

        /// <summary>
        /// Called just before the current solution closes
        /// </summary>
        public void solutionEvents_BeforeClosing()
        {
            // Clear the task list
            TaskViewModel.ClearTaskCollections();
        }

        /// <summary>
        /// Called when the current open window changes in the editor
        /// </summary>
        /// <param name="GotFocus">Window that has focus</param>
        /// <param name="LostFocus">Window that lost focus</param>
        private void windowEvents_WindowActivated(Window GotFocus, Window LostFocus)
        {
            // Call new document opened on control
            RefilterTaskListOnFileChange();
        }

        /// <summary>
        /// Called whenever any line is changed in the current text document
        /// </summary>
        /// <param name="StartPoint"></param>
        /// <param name="EndPoint"></param>
        /// <param name="Hint"></param>
        private void textEditorEvents_LineChanged(TextPoint StartPoint, TextPoint EndPoint, int Hint)
        {
            taskUpdateDelayTimer.Reset();
        }

        #endregion

        #region Timers

        /// <summary>
        /// This event is waiting for the internal task list to update then
        /// updates and filters the add-on list
        /// </summary>
        private void taskUpdateTimer_Elapsed(object source, ElapsedEventArgs e)
        {
            ReadAllTasksAndFilter();
        }

        /// <summary>
        /// When the solution is opened, it can take a long timer before the tasks are read through
        /// This timer will try to reread all the tasks. If there are no tasks found, it will retry
        /// for a certain amount of timer before giving up
        /// </summary>
        private void solutionOpenedTimer_Elapsed(object source, ElapsedEventArgs e)
        {
            ReadAllTasksAndFilter();

            // If the number of tasks found havne't changed and we haven't tried too many times, reset the timer to retry this code
            // The task list count needs to stabilize, then we can stop checking
            if (retrySolutionOpenedCount < 20)
            {
                if (TaskViewModel.TaskCollection.Count == 0 || TaskViewModel.TaskCollection.Count != currentListCount)
                {
                    // Update the local variable
                    currentListCount = TaskViewModel.TaskCollection.Count;
                    // Reset the timer to check again in the timeout time
                    solutionOpenedDelayTimer.Reset();
                    // Increment the count
                    retrySolutionOpenedCount++;
                }
            }
        }

        #endregion
    }
}
