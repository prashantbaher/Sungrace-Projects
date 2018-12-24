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

            //// Sets Solidworks application to swApp 
            //swApp = new SldWorks();

            // Sets current document as active document
            swDoc = (ModelDoc2)swApp.ActiveDoc;

            // Sets the current document to part document
            swPart = (PartDoc)swDoc;

            // Gets the body of part and sets to the variable 
            swBody = (Body2)swPart.Body();

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
                    
                }

            } while (swFace != null);
        }

        #endregion

    }
}
