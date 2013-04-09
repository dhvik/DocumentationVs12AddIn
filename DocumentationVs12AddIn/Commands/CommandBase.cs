using System.Drawing;
using System.Text.RegularExpressions;
using EnvDTE;
using EnvDTE80;

namespace DocumentationVs12AddIn.Commands {
	/// <summary>
	/// Summary description for CommandBase.
	/// </summary>
	/// <remarks>
	/// 2013-04-08 dan: Created
	/// </remarks>
	public abstract class CommandBase {
		/* *******************************************************************
         *  Properties 
         * *******************************************************************/
		#region public Command Command
		/// <summary>
		/// Get/Sets the Command of the CommandBase
		/// </summary>
		/// <value></value>
		//public Command Command { get; set; }
		#endregion
		#region protected DTE2 DTE
		/// <summary>
		/// Get/Sets the DTE of the CommandBase
		/// </summary>
		/// <value></value>
		public DTE2 DTE { get; set; }
		#endregion
		protected TextSelection Selection {
			get { return ((TextSelection) DTE.ActiveDocument.Selection); }
		}
		protected bool IsHtml {
			get { return DTE.ActiveWindow.Document.Language == "HTML"; }
		}
		protected bool IsBasic {
			get { return DTE.ActiveWindow.Document.Language == "Basic"; }
		}
		protected bool IsCSharp {
			get { return DTE.ActiveWindow.Document.Language == "CSharp"; }
		}
		/* *******************************************************************
         *  Methods 
         * *******************************************************************/
		#region public virtual void QueryStatus(vsCommandStatusTextWanted neededText, ref vsCommandStatus status, ref object commandText)
		/// <summary>
		/// 
		/// </summary>
		/// <param name="neededText"></param>
		/// <param name="status"></param>
		/// <param name="commandText"></param>
		public virtual void QueryStatus(vsCommandStatusTextWanted neededText, ref vsCommandStatus status, ref object commandText) {
			status = vsCommandStatus.vsCommandStatusSupported | vsCommandStatus.vsCommandStatusEnabled;
		}
		#endregion
	

		/// <summary>
		/// Gets the point from the supplied textpoint
		/// </summary>
		/// <param name="tp"></param>
		/// <returns></returns>
		protected Point GetPoint(TextPoint tp = null) {
			if (tp == null) {
				tp = Selection.ActivePoint;
			}
			return new Point(tp.LineCharOffset, tp.Line);
		}

		/// <summary>
		///  Moves the current Selection to the supplied point (offset, line) (x,y)
		/// </summary>
		/// <param name="p"></param>
		/// <param name="extend"></param>
		protected void MoveToPoint(Point p, bool extend = false) {
			Selection.MoveToLineAndOffset(p.Y, p.X, extend);
		}

		protected string CutLeftWord() {
			
			//get templateName
			Selection.WordLeft(true);
			string word = Selection.Text;

			//delete templateName
			Selection.Delete();

			return word;
		}

		/// <summary>
		/// Gets an xml documentation see tag
		/// </summary>
		/// <param name="type">Name of the type to refer to</param>
		/// <returns>The see tag</returns>
		protected string GetSeeXmlDocTag(string type) {
			string pre = "";
			string post = "";
			//remove namespace from type.
			//type = Regex.Replace(type, "^.*\.", "", RegexOptions.None)

			Match m = Regex.Match(type, "^(.*)\\s*\\[\\s*\\]\\s*$");
			if (m.Success) {
				pre = "";
				type = m.Groups[1].Value;
				post = " array";
			}
			//convert generics <> to {}
			type = type.Replace("<", "{").Replace(">", "}");

			return pre + "<see cref=\"" + type + "\"/>" + post;
		}

		///<summary>
		///Fixes the selection so that the selection starts at the first column of the first line
		///and the active point is at the bottom. if we have started to mark part of the last line, inlcude the whole line
		///</summary>
		protected TextSelection FixSelection() {

			var sel = Selection;

			//make sure that the selection starts at the first column of the first line
			//and the active point is at the bottom
			if (sel.ActivePoint.EqualTo(sel.BottomPoint)) {
				sel.SwapAnchor();
			}
			sel.StartOfLine(vsStartOfLineOptions.vsStartOfLineOptionsFirstColumn, true);
			sel.SwapAnchor();

			//if we have started to mark part of the last line, inlcude the whole line
			if (sel.ActivePoint.LineCharOffset > 1) {
				sel.EndOfLine(true);
				sel.LineDown(true);
				sel.StartOfLine(vsStartOfLineOptions.vsStartOfLineOptionsFirstColumn, true);
			}
			return sel;
		}
	}
}