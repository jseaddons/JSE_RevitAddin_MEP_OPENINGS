using System;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Models.EventArgs
{
    /// <summary>
    /// Event arguments for profile change events
    /// </summary>
    public class ProfileChangedEventArgs : System.EventArgs
    {
        public UserProfile Profile { get; }
        public DateTime Timestamp { get; }
        public string ChangeType { get; }

        public ProfileChangedEventArgs(UserProfile profile, string changeType = "Changed")
        {
            Profile = profile ?? throw new ArgumentNullException(nameof(profile));
            Timestamp = DateTime.Now;
            ChangeType = changeType;
        }
    }
}
