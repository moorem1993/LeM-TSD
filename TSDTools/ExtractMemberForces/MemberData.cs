using System;
using System.Collections.Generic;
using System.Text;

namespace ExtractMemberForces
{
    class MemberData
    {
		#region Properties

		/// <summary>
		/// Gets or sets the member ID
		/// </summary>
		public string Guid { get; set; }

		/// <summary>
		/// Gets or sets the member name
		/// </summary>
		public string Name { get; set; }

		/// <summary>
		/// Gets or sets the members type
		/// </summary>
		public string Type { get; set; }

		/// <summary>
		/// Gets or sets the members section
		/// </summary>
		public string Section { get; set; }

		/// <summary>
		/// Gets or sets the members material
		/// </summary>
		public string Material { get; set; }

		/// <summary>
		/// Gets or sets the members total length (includes all spans)
		/// </summary>
		public double TotalLength { get; set; }

		/// <summary>
		/// Gets or sets the span within the member
		/// </summary>
		public int Span { get; set; }

		/// <summary>
		/// Gets or sets the span length
		/// </summary>
		public double SpanLength { get; set; }

		/// <summary>
		/// Gets or sets the position along the member at which the forces are taken from (ranges from 0 to 1)
		/// </summary>
		public double Position { get; set; }

		/// <summary>
		/// Gets or sets the loadcase
		/// </summary>
		public string LoadCase { get; set; }

		/// <summary>
		/// Gets or sets Fx
		/// </summary>
		public double Fx { get; set; }

		/// <summary>
		/// Gets or sets Fy
		/// </summary>
		public double Fy { get; set; }

		/// <summary>
		/// Gets or sets Fz
		/// </summary>
		public double Fz { get; set; }

		/// <summary>
		/// Gets or sets Mxx
		/// </summary>
		public double Mxx { get; set; }

		/// <summary>
		/// Gets or sets Myy
		/// </summary>
		public double Myy { get; set; }

		/// <summary>
		/// Gets or sets Mzz
		/// </summary>
		public double Mzz { get; set; }

		#endregion
	}
}
