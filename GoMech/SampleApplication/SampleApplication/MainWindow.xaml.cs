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

            // Loads this function when application start
            StartupMethods();
        }

        #endregion

        #region Private Methods

        /// <summary>
        /// Loads these things when application start
        /// </summary>
        public void StartupMethods()
        {
            // Getting directory information of Example folder
            DirectoryInfo outputDirectoryInfo = new DirectoryInfo(@"C:\ProgramData\Hypertherm CAM\ProNest 2019\Examples\");

            // Assigning the dxf file to cad part list 
            CADPartList.ItemsSource = outputDirectoryInfo.GetFiles("*.dxf"); 
        }

        /// <summary>
        /// Run Autonest function in ProNest
        /// </summary>
        /// <param name="sender">Button itself is the sender</param>
        /// <param name="e">Event arguments for this button</param>
        public void Button_Click(object sender, RoutedEventArgs e)
        {
            SungraceApplication.Visibility = Visibility.Hidden;

            // Object for our ProNest application
            IpnApplication pronestObject;

            // Object for Pronest Program Id
            string pronestProgId = "ProNest.Application.12";

            // Setting type of ProNest by using pronest program id
            var pronestType = Type.GetTypeFromProgID(pronestProgId);

            // Creating instance of ProNest application
            pronestObject = (IpnApplication)Activator.CreateInstance(pronestType);

            // Making ProNest visible for user
            pronestObject.Visible = PronestAppVisibility.IsChecked.Value ? true : false;

            // Getting CAD properties of imported part
            IpnCADImportProperties CADImportProperties = pronestObject.Job.PartList.GetCADImportProperties();

            // Getting Defult nesting properties
            IpnNestingProperties NestingProperties = pronestObject.Job.PartList.GetNestingProperties();

            // Setting number of nest required
            NestingProperties.QuantityRequired = int.Parse(NestQuantityTextBox.Text);

            // Setting material of importing cad file
            // NestingProperties.Material = pronestObject.Job.Machine.Materials.GetMaterialByID(18);
            int selectedMaterial = int.Parse((string)MaterialList.SelectedValue);
            NestingProperties.Material = pronestObject.Job.Machine.Materials.GetMaterialByID(selectedMaterial);


            // Re-applying the leads
            NestingProperties.RetainAllExistingLeads = false;

            // Setting class of part to "40Amp N2/N2" before importing cad file
            NestingProperties.PartClass = MaterialList.Text;

            // Path to our sample dxf files
            // string CADfilePath = @"C:\ProgramData\Hypertherm CAM\ProNest 2019\Examples\Base Plate.dxf";
            string CADfilePath = CADPartList.SelectedValue.ToString();

            // Adding CAD file to partlist
            pronestObject.Job.PartList.AddCADFile(CADfilePath, NestingProperties, CADImportProperties);

            // Getting default setup values for before starting autonest
            IpnAutoNestSetupValues AutoNestSetupValues = pronestObject.Job.AutoNest.GetSetupValues();

            // Assigning cutome plate width and length
            // NOTE: ProNest use Length first then width
            AutoNestSetupValues.CustomPlateWidth = 5;
            AutoNestSetupValues.CustomPlateLength = 3;

            // Starting autonest
            pronestObject.Job.AutoNest.Start(AutoNestSetupValues, false, false);

            // Setting default output folder
            string outputFolder = pronestObject.Job.Machine.Settings.ReadString("NC Parameters", "default output folder", @"C:\ProNest\");

            // Setting default output extension
            string outputExtension = pronestObject.Job.Machine.Settings.ReadString("NC Parameters", "default_cnc_extension", "CNC");

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

            // Setting output index [Right now I don't know why we use it]
            int OutputIndex = 1;

            // Looping through each nest in the list of all nest done in our job
            foreach (IpnNest nest in pronestObject.Job.Nests)
            {
                // Output every nest
                nest.Output(outputFileFolderName + "\\" + pronestObject.Job.Name + "_" + OutputIndex.ToString() + "." + outputExtension, 0, false, true);

                nest.ExportDXF(outputFileFolderName + "\\" + pronestObject.Job.Name + "_" + OutputIndex.ToString() + "." + "DXF");

                // Increasing the index by 1
                OutputIndex++;
                
            }

            // Saving this job
            pronestObject.Job.SaveAs(outputFileFolderName + "\\" + pronestObject.Job.Name + ".nif", null);

            // Output message to user
            MessageBox.Show("Job Successfully Completed", "Sungrace Application", MessageBoxButton.OK);

            // Visible our application after hitting Ok button
            SungraceApplication.Visibility = Visibility.Visible;

            // Quitting the ProNest after completing job
            // pronestObject.Quit();
        }

        #region Up and Down Arrow Buttons

        /// <summary>
        /// Increment the quantity of required nest
        /// </summary>
        /// <param name="sender">Button itself is the sender</param>
        /// <param name="e">Event arguments for this button</param>
        private void UpArrowButton(object sender, RoutedEventArgs e)
        {
            // Getting nest quantity text and storing into integer
            int nestQuantity = int.Parse(NestQuantityTextBox.Text);

            // Increasing the nest quantity by 1
            ++nestQuantity;

            // Setting the incremental quantity inplace of original text
            NestQuantityTextBox.Text = nestQuantity.ToString();
        }

        /// <summary>
        /// Decrement the quantity of required nest
        /// </summary>
        /// <param name="sender">Button itself is the sender</param>
        /// <param name="e">Event arguments for this button</param>
        private void DownArrowButton(object sender, RoutedEventArgs e)
        {
            // Getting nest quantity text and storing into integer
            int nestQuantity = int.Parse(NestQuantityTextBox.Text);

            // Decresing the nest quantity by 1
            --nestQuantity;

            // If nest quantity is less than 0 do not decrement
            if (!(nestQuantity <= 0))
            {
                // Setting the Decremental quantity inplace of original text
                NestQuantityTextBox.Text = nestQuantity.ToString();
            }

        }


        #endregion

        #endregion

        /*
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            string[] StagedFiles = Directory.GetFiles(@"C:\ProNest");

            foreach (string file in StagedFiles)
            {

                textBox1.Text += file + Environment.NewLine;
            }
        }
        */
    }
}
