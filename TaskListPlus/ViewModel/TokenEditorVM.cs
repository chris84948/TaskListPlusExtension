using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using chrisbjohnson.TaskListPlus.MVVMLib;
using System.Text.RegularExpressions;

namespace chrisbjohnson.TaskListPlus
{
    class TokenEditorVM : ObservableObject
    {
        private String _token;
        private Boolean _tokenValid;
        private Boolean _tokenInvalid;

        private String REGEX_MATCH = @"[^A-Za-z0-9_]";

        /// <summary>
        /// Keeps track of the token string
        /// </summary>
        public String Token
        {
            get { return _token; }
            set
            {
                _token = value;
                OnPropertyChanged("Token");

                TokenValid = CheckIfTokenIsValid();
            }
        }

        /// <summary>
        /// True when the token is valid
        /// Must be alphanumerical, or underscore
        /// </summary>
        public Boolean TokenValid
        {
            get { return _tokenValid; }
            set 
            { 
                _tokenValid = value;
                OnPropertyChanged("TokenValid");

                // Token invalid is used for the icon - it shouldn't show when the token is empty
                TokenInvalid = !TokenValid || String.IsNullOrEmpty(Token);
            }
        }

        /// <summary>
        /// Inverse of TokenValid (for GUI)
        /// </summary>
        public Boolean TokenInvalid
        {
            get { return _tokenInvalid; }
            set 
            { 
                _tokenInvalid = value;
                OnPropertyChanged("TokenInvalid");
            }
        }

        /// <summary>
        /// Constructor, creates the regex match object to check against when the toen is typed
        /// </summary>
        public TokenEditorVM()
        {
            Token = "";
            TokenValid = false;

            // Set so the error icon is not showing on open
            TokenInvalid = false;
        }

        private Boolean CheckIfTokenIsValid()
        {
            // Check length, it needs to have at least one character
            if (Token.Length == 0) return false;

            // Check regex pattern against the token - will show a match for any invalid character
            return (Regex.IsMatch(Token, REGEX_MATCH) ? false : true);
        }
    }
}
