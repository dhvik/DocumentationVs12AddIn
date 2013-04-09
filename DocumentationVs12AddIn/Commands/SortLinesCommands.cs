using System;
using System.Text.RegularExpressions;
using EnvDTE;

namespace DocumentationVs12AddIn.Commands {
	/// <summary>
	/// Summary description for SortLinesCommands.
	/// </summary>
	/// <remarks>
	/// 2013-04-09 dan: Created
	/// </remarks>
	public class SortLinesCommands : CommandBase {
		#region "Sub SortLinesAscending()"
		///<summary>
		///Sorts the selected lines in ascending order
		///</summary>
		[Command("Text Editor::Ctrl+Alt+s")]
		public void SortLinesAscending() {
			SortLines();
		}
		#endregion
		#region "Sub SortLinesDescending()"
		///<summary>
		///Sorts the selected lines in descending order
		///</summary>
		[Command("Text Editor::Ctrl+Shift+Alt+s")]
		public void SortLinesDescending() {
			SortLines(true);
		}
		#endregion
		#region "Private Function SortLines(Optional ByVal reverse As Boolean = False)"
		/// <summary>
		/// Sorts the selected lines
		/// </summary>
		/// <param name="reverse" ></param>
		/// <returns ></returns>
		private void SortLines(bool reverse = false) {
			TextSelection sel = FixSelection();
			dynamic text = sel.Text.Substring(0, sel.Text.Length - 2);
			string[] lines = Regex.Split(text, Regex.Escape(Environment.NewLine));

			Array.Sort(lines);
			if (reverse) {
				Array.Reverse(lines);
			}
			sel.Delete();
			sel.Insert(string.Join(Environment.NewLine, lines) + Environment.NewLine, (int)vsInsertFlags.vsInsertFlagsContainNewText);
			sel.SwapAnchor();
		}
		#endregion
	}
}