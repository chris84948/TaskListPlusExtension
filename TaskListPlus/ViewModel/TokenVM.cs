using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using chrisbjohnson.TaskListPlus.MVVMLib;

namespace chrisbjohnson.TaskListPlus
{
    /// <summary>
    /// Class to hold all comment tokens add addition/removal
    /// </summary>
    public class TokenVM : ObservableObject
    {
        private Object[] _tokensFull;
        private Object[] _tokens;
        private EnvDTE80.DTE2 ApplicationObject;
        private EnvDTE.Properties properties;

        /// <summary>
        /// Contains a list of all comment tokens as objects on this instance of Visual Studio
        /// </summary>
        public Object[] TokensFull
        {
            get { return _tokensFull; }
            set
            {
                _tokensFull = value;
            }
        }

        /// <summary>
        /// Contains a slightly adjust readable version of the tokens, removing the :2 at the end
        /// Also adding "ALL" as an option and removing the last option "UnresolvedMergeConflict
        /// </summary>
        public Object[] Tokens
        {
            get { return _tokens; }
            set
            {
                _tokens = value;
                OnPropertyChanged("Tokens");
            }
        }

        /// <summary>
        /// Constructor for initially reading all the comment tokens
        /// </summary>
        /// <param name="ApplicationObject"></param>
        public TokenVM(ref EnvDTE80.DTE2 ApplicationObject)
        {
            // Copy the application reference to a local object
            this.ApplicationObject = ApplicationObject;

            // Get all the comment tokens from the application object
            GetCommentTokens();

            // Get the readable list of comment tokens
            GetReadableCommentTokens();
        }

        /// <summary>
        /// Read all the comment tokens from the application object
        /// </summary>
        public void GetCommentTokens() 
        {
            // Variables to represent the properties collection and each property in the Options dialog box.
            properties = default(EnvDTE.Properties);

            // Represents the Task List Node under the Enviroment node.
            properties = ApplicationObject.get_Properties("Environment", "TaskList");

            // Represents the items in the comment Token list and their priorities (1-3/low-high).
            EnvDTE.Property commentProperties = properties.Item("CommentTokens");
            TokensFull = (Object[])commentProperties.Value;
        }

        public void GetReadableCommentTokens()
        {
            // Create an array of strings based on the list of token objects
            Tokens = new String[TokensFull.Length + 1];

            // Add an item for ALL
            Tokens[0] = "ALL";

            for (int i = 1; i < Tokens.Length; i++)
            {
                // Split the name TODO:2 into two sections and just take the first TODO
                Tokens[i] = TokensFull[i - 1].ToString().Split(new char[] { ':' })[0];
            }
        }

        /// <summary>
        /// Add a token to the list of token and save it
        /// </summary>
        /// <param name="tokenName">Name of the new token</param>
        public void AddToken(String tokenName)
        {
            TokensFull = AddItemToArray(TokensFull, tokenName + ":2");

            // Update the properties with the new token
            properties.Item("CommentTokens").Value = TokensFull;

            // Re-read tokens from master
            GetCommentTokens();
            GetReadableCommentTokens();
        }

        /// <summary>
        /// Remove a token from the list of tokens and save it
        /// </summary>
        /// <param name="tokenName">Token name to remove</param>
        public void RemoveToken(String tokenName)
        {
            TokensFull = RemoveItemFromArray(TokensFull, tokenName);

            // Update the properties with the new token
            properties.Item("CommentTokens").Value = TokensFull;

            // Re-read tokens from master
            GetCommentTokens();
            GetReadableCommentTokens();
        }

        /// <summary>
        /// Add a token to the array of tokens by name
        /// </summary>
        /// <param name="array">List of all tokens</param>
        /// <param name="itemToAdd">String name of the token</param>
        /// <returns>Modified array with token added</returns>
        private Object[] AddItemToArray(Object[] array, String itemToAdd)
        {
            Array.Resize(ref array, array.Length + 1);
            array[array.Length - 1] = itemToAdd;
            return array;
        }

        /// <summary>
        /// Remove a token from the array of tokens by name
        /// </summary>
        /// <param name="array">List of all tokens</param>
        /// <param name="itemToRemove">String name of the token</param>
        /// <returns>Modified array with token removed</returns>
        private Object[] RemoveItemFromArray(Object[] array, String itemToRemove)
        {
            Object[] returnArray = new Object[array.Length - 1];
            int positionCount = 0;

            for (int i = 0; i < array.Length; i++)
            {
                String tempString = Regex.Replace(array[i].ToString(), @":\d", "");
                if (!tempString.Equals(itemToRemove, StringComparison.CurrentCultureIgnoreCase))
                {
                    // Copy all items that don't equal the item to remove into the new array
                    returnArray[positionCount] = array[i];
                    // Increment the position count
                    positionCount++;
                }
            }
            return returnArray;
        }
    }
}
