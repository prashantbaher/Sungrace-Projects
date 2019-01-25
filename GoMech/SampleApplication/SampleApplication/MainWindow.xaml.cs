using ProNest;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Input;

namespace SampleApplication
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        #region Public Properties

        // Object for our ProNest application
        public IpnApplication pronestObject;

        // A List For ProNest Materials
        public List<IpnMaterial> materials;

        // List of Material Classess names
        public List<IpnClassNames> ipnClassNames;

        #endregion

        #region Constructor

        /// <summary>
        /// Default constructor
        /// </summary>
        public MainWindow()
        {
            InitializeComponent();

            // Initialize Pronest to access material from Data base
            InitializeProNest();
            
            // Loads this function when application start
            StartupMethods();

            
        }

        /// <summary>
        /// Initialize Pronest to access material from Data base
        /// </summary>
        private void InitializeProNest()
        {
            // Object for Pronest Program Id
            string pronestProgId = "ProNest.Application.12";

            // Setting type of ProNest by using pronest program id
            var pronestType = Type.GetTypeFromProgID(pronestProgId);

            // Creating instance of ProNest application
            pronestObject = (IpnApplication)Activator.CreateInstance(pronestType);

            // Initializing list of materials
            materials = new List<IpnMaterial>();

            // Looping through each material inside ProNest materials
            foreach (IpnMaterial material in  pronestObject.Database.Materials)
            {
                // Adding each material to list
                materials.Add(material);
            }

            // Assigning itemsource of material list combobox text
            MaterialListComboBox.ItemsSource = materials.Select(material => material.Name +" "+ Math.Round(material.Thickness, 3));

        }

        #endregion

        #region Private Methods

        /// <summary>
        /// Loads these things when application start
        /// </summary>
        public void StartupMethods()
        {
            // Getting directory information of Example folder
            DirectoryInfo sampleCADPartFilesLocation = new DirectoryInfo(@"C:\ProgramData\Hypertherm CAM\ProNest 2019\Examples\");

            // Assigning the dxf file to cad part list 
            CADPartListComboBox.ItemsSource = sampleCADPartFilesLocation.GetFiles("*.dxf"); 
        }

        /// <summary>
        /// Run Autonest function in ProNest
        /// </summary>
        /// <param name="sender">Button itself is the sender</param>
        /// <param name="e">Event arguments for this button</param>
        public void Button_Click(object sender, RoutedEventArgs e)
        {
            // Giving warning message to user if job is empty
            if (OutputJobNameTextBox.Text == string.Empty)
                System.Windows.MessageBox.Show("Empty Job name, using default name for saving file.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);

            // Making applications visibilty to hidden
            SungraceApplication.Visibility = Visibility.Collapsed;
            
            // Making ProNest visible for user
            pronestObject.Visible = PronestAppVisibility.IsChecked.Value ? true : false;

            // Getting CAD properties of imported part
            IpnCADImportProperties CADImportProperties = pronestObject.Job.PartList.GetCADImportProperties();

            // Getting Defult nesting properties
            IpnNestingProperties NestingProperties = pronestObject.Job.PartList.GetNestingProperties();

            // Setting number of nest required
            NestingProperties.QuantityRequired = int.Parse(NestQuantityNumberBox.Text);

            // Setting material of importing cad file
            NestingProperties.Material = materials[MaterialListComboBox.SelectedIndex];

            // Re-applying the leads
            NestingProperties.RetainAllExistingLeads = false;

            // Setting class of part to "40Amp N2/N2" before importing cad file
            NestingProperties.PartClass = MaterialListComboBox.Text;

            // Full Path to our sample dxf files
            string CADfilePath = ((FileInfo)CADPartListComboBox.SelectedItem).FullName;

            // Adding CAD file to partlist
            pronestObject.Job.PartList.AddCADFile(CADfilePath, NestingProperties, CADImportProperties);

            // Getting default setup values for before starting autonest
            IpnAutoNestSetupValues AutoNestSetupValues = pronestObject.Job.AutoNest.GetSetupValues();

            // Assigning cutome plate width and length
            // NOTE: ProNest use Length first then width
            // AutoNestSetupValues.CustomPlateWidth = 5;
            // AutoNestSetupValues.CustomPlateLength = 3;

            // Starting autonest
            pronestObject.Job.AutoNest.Start(AutoNestSetupValues, false, false);

            // Setting default output folder
            // string outputFolder = pronestObject.Job.Machine.Settings.ReadString("NC Parameters", "default output folder", @"C:\ProNest\");
            string outputFolder = pronestObject.Job.Machine.Settings.ReadString("NC Parameters", "default output folder", OutputDirectoryNameTextBox.Text);

            // Setting default output extension
            string outputExtension = pronestObject.Job.Machine.Settings.ReadString("NC Parameters", "default_cnc_extension", "CNC");

            // Getting output job name and assigning it to project job name
            pronestObject.Job.Name = OutputJobNameTextBox.Text;

            // Setting Folder name of output job
            string outputFileFolderName = outputFolder + "\\" + pronestObject.Job.Name;

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
                // If NC file check box is checked then only Output every nest
                if (GenerateNCFileCheckBox.IsChecked == true)
                    nest.Output(outputFileFolderName + "\\" + pronestObject.Job.Name + "_" + OutputIndex.ToString() + "." + outputExtension, 0, false, true);

                // If NC file check box is checked then only export to dxf
                if (ExportToDxfCheckBox.IsChecked == true)
                    nest.ExportDXF(outputFileFolderName + "\\" + pronestObject.Job.Name + "_" + OutputIndex.ToString() + "." + "DXF");

                // If Create Job Summary check box is checked then only create job summary
                if (GenerateJobSummaryCheckBox.IsChecked == true)
                {
                    // Reading template report file
                    string jobSummary = File.ReadAllText(@"Job-Summary\report.html");

                    // Adding Nest value
                    jobSummary = jobSummary.Replace("NestValue", nest.Plate.QuantityNested.ToString());

                    // Adding Time cut value
                    jobSummary = jobSummary.Replace("TimesCutValue", nest.TimesCut.ToString());

                    // Adding Plate utilization value
                    jobSummary = jobSummary.Replace("PlateUsedValue", Math.Round(nest.Costing.PlateUsedUtilization, 3).ToString() + " %");

                    // Adding Plate Material Name
                    jobSummary = jobSummary.Replace("PlatesMaterialValue", nest.Plate.Material.Name);

                    // Adding Nest Dimension
                    jobSummary = jobSummary.Replace("NestDimensionValue", Math.Round(nest.LengthUsed, 3).ToString() + " X " + Math.Round(nest.WidthUsed, 3).ToString() + "in");

                    TimeSpan timeSpan = TimeSpan.FromSeconds(Math.Round(nest.Costing.ProductionTime, 3));
                    //DateTime dateTime = DateTime.Today.Add(timeSpan);
                    string displayTime = timeSpan.ToString();

                    // Adding Production time value
                    jobSummary = jobSummary.Replace("ProductionTimeValue", displayTime);

                    // Adding Production cost value
                    jobSummary = jobSummary.Replace("ProductionCostValue", Math.Round(nest.Costing.ProductionCost, 3).ToString());

                    // Getting job summary file
                    FileInfo fileInfo = new FileInfo(outputFileFolderName + "\\" + "jobSummary.html");

                    // If job summary file is not created then create at output location
                    if (!fileInfo.Exists)
                    {
                        fileInfo.Create().Close();
                    }

                    // Getting StreaWriter for writing new job summary
                    StreamWriter streamWriter = new StreamWriter(fileInfo.FullName);

                    // StreamWrite the new job summary
                    streamWriter.Write(jobSummary);

                    // Closing StreamWriter
                    streamWriter.Close();

                    // Coping logo image for output path
                    File.Copy(@"Job-Summary\logo.png", outputFileFolderName + "\\" + "logo.png", true);
                }

                // Increasing the index by 1
                OutputIndex++;

            }

            // Saving this job if save job check box is checked
            if (SaveJobCheckBox.IsChecked == true)
                pronestObject.Job.SaveAs(outputFileFolderName + "\\" + pronestObject.Job.Name + ".nif", null);

            // Output message to user
            System.Windows.MessageBox.Show("Job Successfully Completed", "Sungrace Application", MessageBoxButton.OK);

            // Visible our application after hitting Ok button
            SungraceApplication.Visibility = Visibility.Visible;

            // Quitting the ProNest after completing job
            if (PronestAppVisibility.IsChecked != true)
                pronestObject.Quit();
        }

        /// <summary>
        /// Browse output directory
        /// </summary>
        /// <param name="sender">Button itself is the sender</param>
        /// <param name="e">Event arguments for this button</param>
        private void BrowseOutputDirectory(object sender, RoutedEventArgs e)
        {
            // Getting folder browsing dialog box
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();

            // Display folder browser
            folderBrowserDialog.ShowDialog();

            // Getting seleted path and assign it to the output directory name textbox
            OutputDirectoryNameTextBox.Text = folderBrowserDialog.SelectedPath;
        }

        #endregion


    }
}
