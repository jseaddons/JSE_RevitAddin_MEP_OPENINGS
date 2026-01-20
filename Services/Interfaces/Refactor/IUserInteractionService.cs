using System.Windows.Forms;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor
{
    /// <summary>
    /// Team J: User interaction abstraction (for testability).
    /// SOLID: Interface Segregation - focused interface for user interactions.
    /// </summary>
    public interface IUserInteractionService
    {
        /// <summary>
        /// Gets filter name from user via input dialog.
        /// </summary>
        string GetFilterNameFromUser(string title, string prompt, string defaultValue);
        
        /// <summary>
        /// Shows error message to user.
        /// </summary>
        void ShowError(string message);
        
        /// <summary>
        /// Shows warning message to user.
        /// </summary>
        void ShowWarning(string message);
        
        /// <summary>
        /// Shows confirmation dialog to user.
        /// </summary>
        bool ShowConfirmation(string message, string title);
    }
}

