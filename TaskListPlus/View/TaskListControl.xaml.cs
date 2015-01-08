using System.Windows.Controls;
using System.Windows.Input;
using EnvDTE80;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Shell;

namespace chrisbjohnson.TaskListPlus
{
    /// <summary>
    /// Interaction logic for TaskListControl.xaml
    /// </summary>
    public partial class TaskListControl : UserControl
    {
        public TaskListControlVM viewModel;

        public TaskListControl()
        {
            InitializeComponent();

            viewModel = new TaskListControlVM();
            this.DataContext = viewModel;
        }

        private void DataGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            // If a double click is registered in blank space, do nothing
            if (dataGridTasks.SelectedIndex == -1) return;

            try
            {
                // Call navigate to to jump to code
                ((TaskListItemVM) dataGridTasks.SelectedItem).NavigateToItem();
                
            }
            catch
            {
                // Do nothing here - if the navigate fails, don't cause an exception
            }
        }
    }
}
