using ProNest;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Windows.Input;

namespace SampleApplication
{
    /// <summary>
    /// Class to hold <see cref="ProNest"/> Application functionality
    /// </summary>
    public class ProNestViewModel : BaseViewModel
    {
        #region Public Properties

        /// <summary>
        /// ProNest application object
        /// </summary>
        public IpnApplication PronestObject { get; set; }

        /// <summary>
        /// ProNest Material for Binding with Material ComboBox
        /// </summary>
        public IEnumerable<string> Materials { get; set; }

        /// <summary>
        /// List to hold all <see cref="ProNest.IpnMaterial"/>
        /// </summary>
        public List<IpnMaterial> PronestMaterials { get; set; } = new List<IpnMaterial>();

        /// <summary>
        /// ProNest Machine for Binding with Machine ComboBox
        /// </summary>
        public IEnumerable<string> Machines { get; set; }

        /// <summary>
        /// List to hold all <see cref="ProNest.IpnMachine"/>
        /// </summary>
        public List<IpnMachine> PronestMachines { get; set; } = new List<IpnMachine>();

        /// <summary>
        /// A check for ProNest visibilty
        /// </summary>
        public bool ProNestVisibilityCheck { get; set; }

        /// <summary>
        /// List to hold all <see cref="ProNest.IpnMaterial.Classes"/>
        /// </summary>
        public List<string> PronestMaterialsClasses { get; set; } = new List<string>() { "None", "40Amp N2/N2", "40Amp Air/Air" };

        /// <summary>
        /// CAD Part List for Binding with CADPartList ComboBox
        /// </summary>
        public List<FileInfo> CadPartList { get; set; } = new List<FileInfo>();

        /// <summary>
        /// Job name in <see cref="PronestObject"/>
        /// </summary>
        public string JobName { get; set; } = "Job1";

        /// <summary>
        /// Nest quantity for Nesting
        /// </summary>
        public int NestQuantity { get; set; } = 1;

        /// <summary>
        /// Default output directory path
        /// </summary>
        public string DefaultOutputDirectoryPath { get; set; } = @"C:\ProNest\";

        /// <summary>
        /// Selection Index for Material
        /// </summary>
        public int MaterialIndex { get; set; }

        /// <summary>
        /// Selection Index for Machine
        /// </summary>
        public int MachineIndex { get; set; }

        /// <summary>
        /// Selection Index for Part file
        /// </summary>
        public int PartFileIndex { get; set; }

        /// <summary>
        /// Selection Index for Class
        /// </summary>
        public int ClassIndex { get; set; }

        /// <summary>
        /// Check if NC file CheckBox is selected
        /// </summary>
        public bool NCFileCheck { get; set; }

        /// <summary>
        /// Check if Export To Dxf CheckBox is selected
        /// </summary>
        public bool ExportToDxfCheck { get; set; }

        /// <summary>
        /// Check if Job summary CheckBox is selected
        /// </summary>
        public bool JobSummaryCheck { get; set; }

        /// <summary>
        /// Check if Save Job CheckBox is selected
        /// </summary>
        public bool SaveJobCheck { get; set; }


        #endregion

        #region Public Commands

        /// <summary>
        /// The command for ProNest output
        /// </summary>
        public ICommand ProNestOutput { get; set; }

        /// <summary>
        /// The command for Browsing Directory
        /// </summary>
        public ICommand BrowseDirectory { get; set; }

        #endregion

        #region Constructor

        /// <summary>
        /// Default constructor
        /// </summary>
        public ProNestViewModel()
        {
            // Initialize Pronest to access material from Data base
            InitializeProNest();

            // Run Browser directory command
            BrowseDirectory = new RelayCommand(() => BrowseOutputDirectory());

            // Run output method to give final output
            ProNestOutput = new RelayCommand(() => OutputMethods());
        }

        #endregion

        #region Helpers

        /// <summary>
        /// Initialize Pronest to access material from Data base
        /// </summary>
        public void InitializeProNest()
        {
            // Program Id of ProNest
            string pronestProgId = "ProNest.Application.12";

            // Gets the program type from Pronest Program Id
            var pronestType = Type.GetTypeFromProgID(pronestProgId);

            // Creates instance of ProNest
            PronestObject = (IpnApplication)Activator.CreateInstance(pronestType);

            // Looping through each material in ProNest Database
            foreach (IpnMaterial material in PronestObject.Database.Materials)
            {
                // Add each material into ProNest material list
                PronestMaterials.Add(material);
            }

            // Binding ProNest material to Material List combobox itemsource
            Materials = PronestMaterials.Select(material => material.Name + " " + Math.Round(material.Thickness, 3));

            // Looping through each machine in ProNest
            foreach (IpnMachine machine in PronestObject.Machines)
            {
                // Add each machine to ProNest Machine list
                PronestMachines.Add(machine);
            }

            // Binding ProNest machine to Machine List combobox itemsource
            Machines = PronestMachines.Select(machine => machine.Name);

            // Getting directory information of Example folder
            DirectoryInfo sampleCADPartFilesLocation = new DirectoryInfo(@"C:\ProgramData\Hypertherm CAM\ProNest 2019\Examples\");

            // Looping through each dxf file in the list of files
            foreach (FileInfo CadFile in sampleCADPartFilesLocation.GetFiles("*.dxf"))
            {
                // Adding each filter file into the CAD Part List
                CadPartList.Add(CadFile);
            }
        }

        /// <summary>
        /// Method for browsing output directory
        /// </summary>
        private void BrowseOutputDirectory()
        {
            // Getting folder browsing dialog box
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();

            // Display folder browser
            folderBrowserDialog.ShowDialog();

            // Getting seleted path and assign it to the output directory name textbox
            DefaultOutputDirectoryPath = folderBrowserDialog.SelectedPath;
        }

        /// <summary>
        /// Final output methods
        /// </summary>
        private void OutputMethods()
        {
            // If ProNest visibility check box is checked then Make ProNest visible
            if (ProNestVisibilityCheck == true)
                PronestObject.Visible = true;

            // Method for Auto-Nesting
            AutoNestProcess();

            // Method for Files output
            OutoutFiles();
        }

        /// <summary>
        /// Method for Auto-Nesting
        /// </summary>
        private void AutoNestProcess()
        {
            // Getting properties of imported CAD files
            IpnCADImportProperties CADImportProperties = PronestObject.Job.PartList.GetCADImportProperties();

            // Getting Defult nesting properties
            IpnNestingProperties NestingProperties = PronestObject.Job.PartList.GetNestingProperties();

            // Setting the nest quantity
            NestingProperties.QuantityRequired = NestQuantity;

            // Setting material of cad files
            // this could throw error
            NestingProperties.Material = PronestMaterials[MaterialIndex];

            // Not retaining existing the leads
            NestingProperties.RetainAllExistingLeads = false;

            // Setting class to cad part
            NestingProperties.PartClass = PronestMaterialsClasses[ClassIndex].ToString();

            // Selected Cad File
            string SelectedCadFile = CadPartList[PartFileIndex].FullName;

            // Adding CAD file to Part list
            PronestObject.Job.PartList.AddCADFile(SelectedCadFile, NestingProperties, CADImportProperties);

            // Getting defult setup value before Auto nesting
            IpnAutoNestSetupValues AutoNestSetupValues = PronestObject.Job.AutoNest.GetSetupValues();

            // Starting Auto-Nest
            PronestObject.Job.AutoNest.Start(AutoNestSetupValues, false, false);
        }

        /// <summary>
        /// Method for Files output
        /// </summary>
        private void OutoutFiles()
        {
            // Setting Output Folder
            string BrowseOutputFolder = DefaultOutputDirectoryPath;

            // Getting browse output folder information
            DirectoryInfo BrowseDirectoryInfo = new DirectoryInfo(BrowseOutputFolder);

            // If browsed output folder name did not exist, then create a new one
            if (!BrowseDirectoryInfo.Exists)
            {
                BrowseDirectoryInfo.Create();
            }

            // Setting default output extension
            string outputExtension = PronestObject.Job.Machine.Settings.ReadString("NC Parameters", "default_cnc_extension", "CNC");

            // Getting output job name and assigning it to project job name
            PronestObject.Job.Name = JobName;

            // Setting Folder name of output job
            string outputFileFolderName = BrowseOutputFolder + "\\" + PronestObject.Job.Name;

            // Getting output foldername information
            DirectoryInfo OutputDirectoryInfo = new DirectoryInfo(outputFileFolderName);

            // If output folder name did not exist, then create a new one with the output folder name
            if (!OutputDirectoryInfo.Exists)
            {
                OutputDirectoryInfo.Create();
            }

            // Setting output index [Right now I don't know why we use it]
            int OutputIndex = 1;

            // Looping through each nest in the list of all nest done in our job
            foreach (IpnNest nest in PronestObject.Job.Nests)
            {
                // If NC file check box is checked then only Output every nest
                if (NCFileCheck == true)
                    nest.Output(outputFileFolderName + "\\" + PronestObject.Job.Name + "_" + OutputIndex.ToString() + "." + outputExtension, 0, false, true);

                // If NC file check box is checked then only export to dxf
                if (ExportToDxfCheck == true)
                    nest.ExportDXF(outputFileFolderName + "\\" + PronestObject.Job.Name + "_" + OutputIndex.ToString() + "." + "DXF");

                // If Create Job Summary check box is checked then only create job summary
                if (JobSummaryCheck == true)
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

                // Saving this job if save job check box is checked
                if (SaveJobCheck == true)
                    PronestObject.Job.SaveAs(outputFileFolderName + "\\" + PronestObject.Job.Name + ".nif", null);
            }

        }

        #endregion

    }
}
