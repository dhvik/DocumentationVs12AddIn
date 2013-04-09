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
		#region public string Description
		/// <summary>
		/// Get/Sets the Description of the CommandAttribute
		/// </summary>
		/// <value></value>
		public string Description { get; set; }
		#endregion
		#region public int IconId
		/// <summary>
		/// Get/Sets the IconId of the CommandAttribute
		/// </summary>
		/// <value></value>
		public int IconId { get; set; }
		#endregion
		#region public string KeyBinding
		/// <summary>
		/// Get/Sets the KeyBinding of the CommandAttribute
		/// </summary>
		/// <value></value>
		public string KeyBinding { get; set; }
		#endregion
		#region public string CommandName
		/// <summary>
		/// Get/Sets the CommandName of the CommandAttribute
		/// </summary>
		/// <value></value>
		public string CommandName { get; set; }
		#endregion

		/* *******************************************************************
         *  Constructors 
         * *******************************************************************/
		#region public CommandAttribute()
		/// <summary>
		/// Initializes a new instance of the <b>CommandAttribute</b> class.
		/// </summary>
		public CommandAttribute() { }
		#endregion
		#region public CommandAttribute(string name, string keyBinding)
		/// <summary>
		/// Initializes a new instance of the <b>CommandAttribute</b> class.
		/// </summary>
		/// <param name="name"></param>
		/// <param name="keyBinding"></param>
		public CommandAttribute(string name, string keyBinding) {
			Name = name;
			KeyBinding = keyBinding;
			Description = Name;
			IconId = 1000;
		}
		#endregion
		#region public CommandAttribute(string keyBinding)
		/// <summary>
		/// Initializes a new instance of the <b>CommandAttribute</b> class.
		/// </summary>
		/// <param name="keyBinding"></param>
		public CommandAttribute(string keyBinding) : this(null, keyBinding) { }
		#endregion
	}
}