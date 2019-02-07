using PropertyChanged;
using System.ComponentModel;

namespace SampleApplication
{
    /// <summary>
    /// A base view model that fires <see cref="PropertyChanged"/> event as needed
    /// </summary>
    [AddINotifyPropertyChangedInterface]
    public class BaseViewModel : INotifyPropertyChanged
    {
        #region Public Events and Methods

        /// <summary>
        /// The event that is fired when any Child property changed
        /// </summary>
        public event PropertyChangedEventHandler PropertyChanged = (sender, e) => { };

        /// <summary>
        /// Call this to fire a <see cref="PropertyChanged"/> event
        /// </summary>
        /// <param name="name">Name of the event to fire</param>
        public void OnPropertyChanged(string name)
        {
            PropertyChanged(this, new PropertyChangedEventArgs(name));
        } 

        #endregion
    }
}
