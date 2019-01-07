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
        /// Soldiworks Loop variable to use <see cref="Loop2"/> interface of Solidworks
        /// </summary>
        public Loop2 swLoop;

        /// <summary>
        /// Solidworks Curve variable to use <see cref="Curve"/> interface of Solidworks
        /// </summary>
        public Curve swCurve;

        /// <summary>
        /// Different surface variables to use <see cref="Surface"/> interface of Solidworks
        /// </summary>
        public Surface swSurface1, swSurface2, swFinalSurface;

        /// <summary>
        /// Different <see cref="Entity"/>
        /// </summary>
        public Entity swEntity, swSafeEntity;

        /// <summary>
        /// EdgeList : List of required Edges in the body
        /// CurveList : List of required Curves in the body
        /// SurfaceList : List of required Surfaces in the body
        /// Edges : List of all edges in the body
        /// EdgeParameter1, EdgeParameter2 : List of Pararmeters of an Edge
        /// Edges : List of faces two faces adjacent to an edge
        /// </summary>
        public ArrayList EdgeList, CurveList, SurfaceList, Edges, EdgeParameter1, EdgeParameter2, Faces;

        /// <summary>
        /// List of integers for iterating
        /// </summary>
        public int i, j, k;

        #endregion

        #region Public Methods

        /// <summary>
        /// Main method
        /// </summary>
        public void Main()
        {
            // Sets new arraylist for edge
            EdgeList = new ArrayList();

            // Sets new arraylist for curves
            CurveList = new ArrayList();

            // Sets new arraylist for Surfaces
            SurfaceList = new ArrayList();

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

            // Gets the body of part and sets to the variable 
            swBody = (Body2)swPart.Body();

            // Check if there are bodies in model or not, if not give a message and end the program
            if (swBody == null)
            {
                swApp.SendMsgToUser("There are no bodies available for this document.");
                return;
            }

            // Gets the first face of the body and assign it to the variable
            swFace = (Face2)swBody.GetFirstFace();

            // Loop through unill faces of a body is nothing
            do
            {
                // Gets the first loop of face and assign it to the variable
                swLoop = (Loop2)swFace.GetFirstLoop();

                // Loop through unill loop of a face is nothing
                while (swLoop != null)
                {
                    swApp.SendMsgToUser("Ok I am ");
                }

            } while (swFace != null);
        }

        #endregion

    }
}
