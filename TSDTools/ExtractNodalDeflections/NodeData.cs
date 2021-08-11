using System;
using System.Collections.Generic;
using System.Text;

namespace ExtractNodalDeflections
{
    class NodeData
    {
        #region properties

        /// <summary>
        /// Gets or sets the node index
        /// </summary>
        public int NodeIndex { get; set; }

        /// <summary>
        /// Gets or sets the loadcase
        /// </summary>
        public string LoadCase { get; set; }

        /// <summary>
        /// Gets or sets movement in the x axis
        /// </summary>
        public double Mx { get; set; }

        /// <summary>
        /// Gets or sets movement in the x axis
        /// </summary>
        public double My { get; set; }

        /// <summary>
        /// Gets or sets movement in the x axis
        /// </summary>
        public double Mz { get; set; }

        /// <summary>
        /// Gets or sets movement in the x axis
        /// </summary>
        public double Rx { get; set; }

        /// <summary>
        /// Gets or sets movement in the x axis
        /// </summary>
        public double Ry { get; set; }

        /// <summary>
        /// Gets or sets movement in the x axis
        /// </summary>
        public double Rz { get; set; }

        #endregion
    }
}
