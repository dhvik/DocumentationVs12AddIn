using System;

namespace DocumentationVs12AddIn.Commands {
	/// <summary>
	/// Summary description for CommandAttribute.
	/// </summary>
	/// <remarks>
	/// 2013-04-08 dan: Created
	/// </remarks>
	[AttributeUsage(AttributeTargets.Method, AllowMultiple = false, Inherited = false)]
	public sealed class CommandAttribute : Attribute {
		/* *******************************************************************
         *  Properties 
         * *******************************************************************/
		#region public string Name
		/// <summary>
		/// Get/Sets the Name of the CommandAttribute
		/// </summary>
		/// <value></value>
		public string Name { get; set; }
		#endregion
		public string Description { get; set; }
		
		public int IconId { get; set; }

		public string KeyBinding { get; set; }

		public string CommandName { get; set; }

		/* *******************************************************************
         *  Constructors 
         * *******************************************************************/

		#region public CommandAttribute()

		/// <summary>
		/// Initializes a new instance of the <b>CommandAttribute</b> class.
		/// </summary>
		public CommandAttribute() { }

		#endregion
		public CommandAttribute(string name, string keyBinding) {
			Name = name;
			KeyBinding = keyBinding;
			Description = Name;
			IconId = 1000;
		}
		public CommandAttribute(string keyBinding):this(null,keyBinding){}
		/* *******************************************************************
         *  Methods 
         * *******************************************************************/

		/* *******************************************************************
		 *  Event methods 
		 * *******************************************************************/
	}
}