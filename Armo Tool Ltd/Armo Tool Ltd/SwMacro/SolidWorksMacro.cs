using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Runtime.InteropServices;
using System;
using System.Collections;
using System.Collections.Generic;

namespace Armo_Tool_Ltd.csproj
{
    /// <summary>
    /// Solidworks Macro for counting Holes
    /// </summary>
    public partial class SolidWorksMacro
    {
        #region Public Variables

        /// <summary>
        ///  The SldWorks swApp variable is pre-assigned for you.
        /// </summary>
        public SldWorks swApp;

        /// <summary>
        /// The Solidworks document
        /// </summary>
        public ModelDoc2 swDoc;

        /// <summary>
        /// Solidworks Part document
        /// </summary>
        public PartDoc swPart;

        /// <summary>
        /// Solidworks Body as <see cref="Body2"/> for our part document of Solidworks
        /// </summary>
        public Body2 swBody;

        /// <summary>
        /// Solidworks face variable to use <see cref="Face2"/> interface of Solidworks
        /// </summary>
        public Face2 swFace;

        /// <summary>
        /// Solidworks surface variables to use <see cref="Surface"/> interface of Solidworks
        /// </summary>
        public Surface swSurface;

        /// <summary>
        /// Solidworks feature variables to use <see cref="Feature"/> interface of Solidworks
        /// </summary>
        public Feature swFeature;

        /// <summary>
        /// Solidworks swHoleWizardData variables to use <see cref="WizardHoleFeatureData2"/> interface of Solidworks
        /// </summary>
        public WizardHoleFeatureData2 swHoleWizardData;

        #endregion

        #region Public Methods

        /// <summary>
        /// Main method
        /// </summary>
        public void Main()
        {
            #region Initial Checks and starting process

            // Sets current document as active document
            swDoc = (ModelDoc2)swApp.ActiveDoc;

            // Check if we opened a file or not and if drawing document is opened
            if ((swDoc == null))
            {
                // Send message to user to open an document to run this add-in
                swApp.SendMsgToUser2("No document found. Please open a document.", (int)swMessageBoxIcon_e.swMbWarning, (int)swMessageBoxBtn_e.swMbOkCancel);
                return;
            }

            // Getting opened document type
            int swDocType = swDoc.GetType();

            // If swDoc is not null then check if it is a drawing, if yes then ask to open assembly or part
            if (swDocType == (int)swDocumentTypes_e.swDocDRAWING)
            {
                swApp.SendMsgToUser("Please open an assembly or a part file");
                return;
            }

            // Sets the current document to part document
            swPart = (PartDoc)swDoc;

            #endregion

            #region Counting Total Cylindrical Surfaces

            // List of bodies in part
            object[] BodyArray = null;

            // Gets all bodies inside a part
            BodyArray = (object[])(swPart.GetBodies2((int)swBodyType_e.swAllBodies, true));

            // Check if there are bodies in model or not, if not give a message and end the program
            if (BodyArray == null)
            {
                // Send error message to user about no body was avaialable
                swApp.SendMsgToUser("There are no bodies available for this document.");
                return;
            }

            // variable for counting number of cylinder in a body
            int swCylinderCount = 0;

            // Looping thourgh each body inside body list
            foreach (Body2 SingleBody in BodyArray)
            {
                // taking each body at a time
                swBody = (Body2)SingleBody;

                // Gets the first face of the body 
                swFace = (Face2)swBody.GetFirstFace();

                // Loop through unill faces of a body is nothing
                do
                {
                    // Getting each surface of this body
                    swSurface = (Surface)swFace.GetSurface();

                    // Check if there are surface in body or not, if not give a message and end the program
                    if (swSurface == null)
                    {
                        // Send error message to user about no surface was avaialable
                        swApp.SendMsgToUser("There are no bodies available for this document.");
                        return;
                    }

                    // If a surface is cylinderical, increase the count of it by 1
                    if (swSurface.IsCylinder())
                        swCylinderCount += 1;

                    // Getting next surface of this body
                    swFace = (Face2)(swFace.GetNextFace());


                } while (swFace != null);
            }

            #endregion

            #region Number of Holes through Hole wizard

            // Getting 1st feature of part document
            swFeature = (Feature)swPart.FirstFeature();

            // Checking if part has some feature, it is overkill to checking such things but it makes our application more full proof.
            if (swFeature == null)
            {
                // Send message to user to open an document to run this add-in
                swApp.SendMsgToUser2("No feature found. Please open a document.", (int)swMessageBoxIcon_e.swMbWarning, (int)swMessageBoxBtn_e.swMbOkCancel);
                return;
            }

            // Variable for counting hole throuhg hole wizard
            int NumberOfWizardHole = 0;

            // Looping though each feature to get wizard hole feature
            do
            {
                // Variable for storing feature name
                string FeatureName = string.Empty;

                // Getting name of each feature inside a part document
                FeatureName = swFeature.GetTypeName2();

                // Checking if feature has some name
                if (FeatureName == null)
                {
                    // Send message to user to open an document to run this add-in
                    swApp.SendMsgToUser2("Something went wrong, Please contact add-in developer to resolve this issue.", (int)swMessageBoxIcon_e.swMbWarning, (int)swMessageBoxBtn_e.swMbOkCancel);
                    return;
                }

                // If feature name is "HoleWzd" (<-- This is the returen value when we apply GetTypeName2 method)
                // Then only count sketch points inside this hole
                if (FeatureName == "HoleWzd")
                {
                    // Getting the defination of feature i.e. What type of feature it is actually
                    swHoleWizardData = (WizardHoleFeatureData2)swFeature.GetDefinition();

                    // Getting the count of sketch point inside hole
                    // because number of sketch point = number of holes through this feature
                    NumberOfWizardHole = swHoleWizardData.GetSketchPointCount();
                }

                // Get and set the next feature inside the document
                swFeature = (Feature)swFeature.GetNextFeature();

            } while (swFeature != null);

            #endregion

            #region Number of actual holes in the document

            // Variable for counting holes
            int HoleCount = 0;

            // Counting actual holes by subtracting number of wizard holes from total number of cylindrical surfaces
            HoleCount = swCylinderCount - NumberOfWizardHole;

            // Message to user about number of holes
            swApp.SendMsgToUser("Number of holes in this document: " + HoleCount);

            #endregion

        }

        #endregion

    }
}
