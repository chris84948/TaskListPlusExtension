using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Shell;
using System.ComponentModel.Design;

namespace chrisbjohnson.TaskListPlus
{
    /// <summary>
    /// This class implements the tool window exposed by this package and hosts a user control.
    ///
    /// In Visual Studio tool windows are composed of a frame (implemented by the shell) and a pane, 
    /// usually implemented by the package implementer.
    ///
    /// This class derives from the ToolWindowPane class provided from the MPF in order to use its 
    /// implementation of the IVsUIElementPane interface.
    /// </summary>
    [Guid("c8bc3106-258b-4638-984c-db2763a9ca9b")]
    public class MyToolWindow : ToolWindowPane
    {
        private TaskListControl control;

        /// <summary>
        /// Standard constructor for the tool window.
        /// </summary>
        public MyToolWindow() :
            base(null)
        {
            this.ToolBar = new CommandID(GuidList.guidTaskListPlusCmdSet, PkgCmdIDList.TaskListPlusToolbar);
            this.ToolBarLocation = (int)VSTWT_LOCATION.VSTWT_TOP;

            // Set the window title reading it from the resources.
            this.Caption = Resources.ToolWindowTitle;
            // Set the image that will appear on the tab of the window frame
            // when docked with an other window
            // The resource ID correspond to the one defined in the resx file
            // while the Index is the offset in the bitmap strip. Each image in
            // the strip being 16x16.
            this.BitmapResourceID = 301;
            this.BitmapIndex = 1;

            // Create the handlers for the toolbar commands.
            var mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;

            if (mcs != null)
            {
                // Scope combo
                // Command for the combo itself
                var scopeComboCmdID = new CommandID(GuidList.guidTaskListPlusCmdSet, (int)PkgCmdIDList.cmdidScope);
                var menuScopeComboCmd = new OleMenuCommand(new EventHandler(ScopeComboHandler), scopeComboCmdID);
                mcs.AddCommand(menuScopeComboCmd);

                // Command for the combo's list
                var scopeComboListCmdID = new CommandID(GuidList.guidTaskListPlusCmdSet, PkgCmdIDList.cmdidScopeList);
                var menuScopeComboListCmdID = new OleMenuCommand(new EventHandler(ScopeComboListHandler), scopeComboListCmdID);
                mcs.AddCommand(menuScopeComboListCmdID);

                // Token combo
                // Command for the combo itself
                var tokenComboCmdID = new CommandID(GuidList.guidTaskListPlusCmdSet, (int)PkgCmdIDList.cmdidToken);
                var menuTokenComboCmd = new OleMenuCommand(new EventHandler(TokenComboHandler), tokenComboCmdID);
                mcs.AddCommand(menuTokenComboCmd);

                // Command for the combo's list
                var tokenComboListCmdID = new CommandID(GuidList.guidTaskListPlusCmdSet, PkgCmdIDList.cmdidTokenList);
                var menuTokenComboListCmdID = new OleMenuCommand(new EventHandler(TokenComboListHandler), tokenComboListCmdID);
                mcs.AddCommand(menuTokenComboListCmdID);

                // Add Token Button
                CommandID addTokenID = new CommandID(GuidList.guidTaskListPlusCmdSet, (int)PkgCmdIDList.cmdidAddToken);
                OleMenuCommand addTokenMenuItem = new OleMenuCommand(AddTokenButtonCallback, addTokenID);
                mcs.AddCommand(addTokenMenuItem);

                // Remove token button
                CommandID removeTokenID = new CommandID(GuidList.guidTaskListPlusCmdSet, (int)PkgCmdIDList.cmdidRemoveToken);
                OleMenuCommand removeTokenMenuItem = new OleMenuCommand(RemoveTokenButtonCallback, removeTokenID);
                mcs.AddCommand(removeTokenMenuItem);

            }

            // This is the user control hosted by the tool window; Note that, even if this class implements IDisposable,
            // we are not calling Dispose on this object. This is because ToolWindowPane calls Dispose on 
            // the object returned by the Content property.
            control = new TaskListControl();
            base.Content = control;
        }

        private void ScopeComboHandler(object sender, EventArgs e)
        {
            if (e == EventArgs.Empty)
            {
                // We should never get here; EventArgs are required.
                throw (new ArgumentException(Resources.EventArgsRequired)); // force an exception to be thrown
            }

            OleMenuCmdEventArgs eventArgs = e as OleMenuCmdEventArgs;

            if (eventArgs != null)
            {
                string newChoice = eventArgs.InValue as string;
                IntPtr vOut = eventArgs.OutValue;

                if (vOut != IntPtr.Zero && newChoice != null)
                {
                    throw (new ArgumentException(Resources.BothInOutParamsIllegal)); // force an exception to be thrown
                }
                else if (vOut != IntPtr.Zero)
                {
                    // when vOut is non-NULL, the IDE is requesting the current value for the combo
                    //if (control != null)
                        Marshal.GetNativeVariantForObject(TaskListControlVM.scopeComboChoices[(int)control.viewModel.SelectedScopeIndex], vOut);
                    //else
                    //    Marshal.GetNativeVariantForObject(TaskListControlVM.scopeComboChoices[(int)TaskListControlVM.taskScope.Solution], vOut);
                }

                else if (newChoice != null)
                {
                    // new value was selected or typed in
                    // see if it is one of our items
                    bool validInput = false;
                    int indexInput = -1;
                    for (indexInput = 0; indexInput < TaskListControlVM.scopeComboChoices.Length; indexInput++)
                    {
                        if (String.Compare(TaskListControlVM.scopeComboChoices[indexInput], newChoice, StringComparison.CurrentCultureIgnoreCase) == 0)
                        {
                            validInput = true;
                            break;
                        }
                    }

                    if (validInput)
                    {
                        control.viewModel.SelectedScopeIndex = indexInput;
                        //ShowMessage(Resources.MyDropDownCombo, this.currentDropDownComboChoice);
                    }
                    else
                    {
                        throw (new ArgumentException(Resources.ParamNotValidStringInList)); // force an exception to be thrown
                    }
                }
                else
                {
                    // We should never get here
                    throw (new ArgumentException(Resources.InOutParamCantBeNULL)); // force an exception to be thrown
                }
            }
            else
            {
                // We should never get here; EventArgs are required.
                throw (new ArgumentException(Resources.EventArgsRequired)); // force an exception to be thrown
            }
        }

        private void ScopeComboListHandler(object sender, EventArgs e)
        {
            if ((null == e) || (e == EventArgs.Empty))
            {
                // We should never get here; EventArgs are required.
                throw (new ArgumentNullException(Resources.EventArgsRequired)); // force an exception to be thrown
            }

            OleMenuCmdEventArgs eventArgs = e as OleMenuCmdEventArgs;

            if (eventArgs != null)
            {
                object inParam = eventArgs.InValue;
                IntPtr vOut = eventArgs.OutValue;

                if (inParam != null)
                {
                    throw (new ArgumentException(Resources.InParamIllegal)); // force an exception to be thrown
                }
                else if (vOut != IntPtr.Zero)
                {
                    Marshal.GetNativeVariantForObject(TaskListControlVM.scopeComboChoices, vOut);
                }
                else
                {
                    throw (new ArgumentException(Resources.OutParamRequired)); // force an exception to be thrown
                }
            }
        }

        private void TokenComboHandler(object sender, EventArgs e)
        {
            if (e == EventArgs.Empty)
            {
                // We should never get here; EventArgs are required.
                throw (new ArgumentException(Resources.EventArgsRequired)); // force an exception to be thrown
            }

            OleMenuCmdEventArgs eventArgs = e as OleMenuCmdEventArgs;

            if (eventArgs != null)
            {
                string newChoice = eventArgs.InValue as string;
                IntPtr vOut = eventArgs.OutValue;

                if (vOut != IntPtr.Zero && newChoice != null)
                {
                    throw (new ArgumentException(Resources.BothInOutParamsIllegal)); // force an exception to be thrown
                }
                else if (vOut != IntPtr.Zero)
                {
                    // when vOut is non-NULL, the IDE is requesting the current value for the combo
                    Marshal.GetNativeVariantForObject(control.viewModel.SelectedToken, vOut);
                }

                else if (newChoice != null)
                {
                    // new value was selected or typed in
                    // see if it is one of our items
                    bool validInput = false;
                    int indexInput = -1;
                    for (indexInput = 0; indexInput < control.viewModel.TokenViewModel.Tokens.Length; indexInput++)
                    {
                        if (String.Compare(control.viewModel.TokenViewModel.Tokens[indexInput].ToString(), newChoice, StringComparison.CurrentCultureIgnoreCase) == 0)
                        {
                            validInput = true;
                            break;
                        }
                    }

                    if (validInput)
                    {
                        control.viewModel.SelectedToken = newChoice;
                        //ShowMessage(Resources.MyDropDownCombo, this.currentDropDownComboChoice);
                    }
                    else
                    {
                        throw (new ArgumentException(Resources.ParamNotValidStringInList)); // force an exception to be thrown
                    }
                }
                else
                {
                    // We should never get here
                    throw (new ArgumentException(Resources.InOutParamCantBeNULL)); // force an exception to be thrown
                }
            }
            else
            {
                // We should never get here; EventArgs are required.
                throw (new ArgumentException(Resources.EventArgsRequired)); // force an exception to be thrown
            }
        }

        private void TokenComboListHandler(object sender, EventArgs e)
        {
            if ((null == e) || (e == EventArgs.Empty))
            {
                // We should never get here; EventArgs are required.
                throw (new ArgumentNullException(Resources.EventArgsRequired)); // force an exception to be thrown
            }

            OleMenuCmdEventArgs eventArgs = e as OleMenuCmdEventArgs;

            if (eventArgs != null)
            {
                object inParam = eventArgs.InValue;
                IntPtr vOut = eventArgs.OutValue;

                if (inParam != null)
                {
                    throw (new ArgumentException(Resources.InParamIllegal)); // force an exception to be thrown
                }
                else if (vOut != IntPtr.Zero)
                {
                    Marshal.GetNativeVariantForObject(control.viewModel.TokenViewModel.Tokens, vOut);
                }
                else
                {
                    throw (new ArgumentException(Resources.OutParamRequired)); // force an exception to be thrown
                }
            }
        }

        private void AddTokenButtonCallback(object sender, EventArgs e)
        {
            // Open the token editor window
            TokenEditor tokenEditor = new TokenEditor();
            tokenEditor.ShowDialog();

            // If the OK button was pressed, add the token to the list
            if (tokenEditor.DialogResult == true)
            {
                control.viewModel.TokenViewModel.AddToken(tokenEditor.Token);
            }

            // Select the new token from the list
            control.viewModel.SelectedToken = tokenEditor.Token;
        }

        private void RemoveTokenButtonCallback(object sender, EventArgs e)
        {
            // Add the new token to the list and control
            control.viewModel.RemoveToken();
        }
    }
}
