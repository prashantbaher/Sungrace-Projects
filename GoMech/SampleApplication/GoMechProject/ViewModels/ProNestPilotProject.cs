using ProNest;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Input;

namespace SampleApplication
{
    /// <summary>
    /// Class to hold <see cref="ProNest"/> Application functionality
    /// </summary>
    public class ProNestPilotProject : BaseViewModel
    {

        #region Public Properties

        /// <summary>
        /// ProNest application object
        /// </summary>
        public IpnApplication pronestObject { get; set; }

        /// <summary>
        /// List to hold all <see cref="ProNest.IpnMaterials"/>
        /// </summary>
        public List<IpnMaterial> pronestMaterials { get; set; } = new List<IpnMaterial>();

        /// <summary>
        /// ProNest Material for Binding with Material ComboBox
        /// </summary>
        public IEnumerable<string> Material { get; set; }

        /// <summary>
        /// CAD Part List for Binding with CADPartList ComboBox
        /// </summary>
        public List<FileInfo> CadPartList { get; set; } = new List<FileInfo>();

        #endregion

        #region Public Commands

        /// <summary>
        /// The command for ProNest visibilty
        /// </summary>
        public ICommand ProNestVisibility { get; set; }

        #endregion

        #region Constructor

        /// <summary>
        /// Default constructor
        /// </summary>
        public ProNestPilotProject()
        {
            // Initialize Pronest to access material from Data base
            InitializeProNest();

            ProNestVisibility = new RelayCommand(() => ShowProNest());
        }

        #endregion

        #region Helpers

        /// <summary>
        /// ProNest visibility action
        /// </summary>
        /// <returns></returns>
        private void ShowProNest()
        {
            // Making pronest visible
            pronestObject.Visible = true;
        }

        /// <summary>
        /// Initialize Pronest to access material from Data base
        /// </summary>
        private void InitializeProNest()
        {
            // Program Id of ProNest
            string pronestProgId = "ProNest.Application.12";

            // Gets the program type from Pronest Program Id
            var pronestType = Type.GetTypeFromProgID(pronestProgId);

            // Creates instance of ProNest
            pronestObject = (IpnApplication)Activator.CreateInstance(pronestType);

            // Looping through each material in ProNest Database
            foreach (IpnMaterial material in pronestObject.Database.Materials)
            {
                // Add each material into ProNest material list
                pronestMaterials.Add(material);
            }

            // Binding ProNest material to Material List combobox itemsource
            Material = pronestMaterials.Select(material => material.Name + " " + Math.Round(material.Thickness, 3));

            // Getting directory information of Example folder
            DirectoryInfo sampleCADPartFilesLocation = new DirectoryInfo(@"C:\ProgramData\Hypertherm CAM\ProNest 2019\Examples\");

            // Looping through each dxf file in the list of files
            foreach (FileInfo CadFile in sampleCADPartFilesLocation.GetFiles("*.dxf"))
            {
                // Adding each filter file into the CAD Part List
                CadPartList.Add(CadFile);
            }
        }

        #endregion

    }
}
