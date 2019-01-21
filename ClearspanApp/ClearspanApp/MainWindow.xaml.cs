using System.Windows;
using SldWorks;
using SwConst;

namespace ClearspanApp
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        #region Default constructor

        /// <summary>
        /// Default constructor
        /// </summary>
        public MainWindow()
        {
            InitializeComponent();
        }

        #endregion

        #region Private Methods

        /// <summary>
        /// Method for creating Tech doc sketch
        /// </summary>
        /// <param name="sender">Sender of this method</param>
        /// <param name="e">Events arguments for this method</param>
        private void CreateTechDocPartSketch(object sender, RoutedEventArgs e)
        {
            #region local variables

            // Solidworks application variable
            SldWorks.SldWorks swApp = new SldWorks.SldWorks();

            // Variable for storing default part template name
            string TemplateName = swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);

            // Create new part document
            ModelDoc2 swDoc = swApp.INewDocument2(TemplateName, 0, 0, 0);

            #endregion

            #region Finishing values
            
            // Setting up Solidworks window state to maximize
            swApp.FrameState = (int)swWindowState_e.swWindowMaximized;

            // Making solidworks visible after completing process
            swApp.Visible = true;

            #endregion
        }

        #endregion
    }
}
