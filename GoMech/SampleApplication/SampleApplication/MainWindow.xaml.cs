using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using ProNest;

namespace SampleApplication
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        #region Constructor

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
        /// Run Autonest function in ProNest
        /// </summary>
        /// <param name="sender">Button itself is the sender</param>
        /// <param name="e">Event arguments for this button</param>
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            // Object for our ProNest application
            IpnApplication pronestObject;

            // Object for Pronest Program Id
            string pronestProgId = "ProNest.Application.12";

            // Setting type of ProNest by using pronest program id
            var pronestType = Type.GetTypeFromProgID(pronestProgId);

            // Creating instance of ProNest application
            pronestObject = (IpnApplication)Activator.CreateInstance(pronestType);

            pronestObject.Visible = true;

            // Getting CAD properties of imported part
            IpnCADImportProperties CADImportProperties = pronestObject.Job.PartList.GetCADImportProperties();

            // Getting Defult nesting properties
            IpnNestingProperties NestingProperties = pronestObject.Job.PartList.GetNestingProperties();

            // Setting number of nest required
            //int nestQuantity = 0;
            NestingProperties.QuantityRequired = int.Parse(NestQuantityTextBox.Text);

            // Setting material of importing cad file
            NestingProperties.Material = pronestObject.Job.Machine.Materials.GetMaterialByID(18);

            // Re-applying the leads
            NestingProperties.RetainAllExistingLeads = false;

            // Setting class of part to "40Amp N2/N2" before importing cad file
            NestingProperties.PartClass = "40Amp N2/N2";

            // Path to our sample dxf file
            string CADfilePath = @"C:\ProgramData\Hypertherm CAM\ProNest 2019\Examples\Base Plate.dxf";

            // Adding CAD file to partlist
            pronestObject.Job.PartList.AddCADFile(CADfilePath, NestingProperties, CADImportProperties);

            // Getting default setup values for before starting autonest
            IpnAutoNestSetupValues AutoNestSetupValues = pronestObject.Job.AutoNest.GetSetupValues();

            // Starting autonest
            pronestObject.Job.AutoNest.Start(AutoNestSetupValues, false, false);

            // Setting default output folder
            string outputFolder = pronestObject.Job.Machine.Settings.ReadString("NC Parameters", "default output folder", @"C:\ProNest\");

            // Setting default output extension
            string outputExtension = pronestObject.Job.Machine.Settings.ReadString("NC Parameters", "default_cnc_extension", "CNC");

            // Setting output index [Right now I don't know why we use it]
            int OutputIndex = 1;

            // Looping through each nest in the list of all nest done in our job
            foreach (IpnNest nest in pronestObject.Job.Nests)
            {
                // Getting output job name and assigning it to project job name
                pronestObject.Job.Name = OutputJobNameTextBox.Text;

                // Setting Folder name of output job
                string outputFileFolderName = outputFolder + pronestObject.Job.Name;

                // Getting output foldername information
                DirectoryInfo directoryInfo = new DirectoryInfo(outputFileFolderName);

                // If output folder name did not exist, then create a new one with the output folder name
                if (!directoryInfo.Exists)
                {
                    directoryInfo.Create();
                }

                // Output every nest
                nest.Output(outputFileFolderName + "\\" + pronestObject.Job.Name + "_" + OutputIndex.ToString() + "." + outputExtension, 0, false, true);

                // Increasing the index by 1
                OutputIndex++;
            }

            // Quitting the ProNest after completing job
            // pronestObject.Quit();


        }

        #endregion
    }
}
