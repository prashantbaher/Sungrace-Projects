using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Runtime.InteropServices;
using System;

namespace Armo_Tool_Ltd.csproj
{
    public partial class SolidWorksMacro
    {
        #region Public Properties

        /// <summary>
        ///  The SldWorks swApp variable is pre-assigned for you.
        /// </summary>
        public SldWorks swApp;

        #endregion

        #region Public Methods

        /// <summary>
        /// Main method
        /// </summary>
        public void Main()
        {
            swApp.SendMsgToUser("Hello world");
        }

        #endregion

    }
}


