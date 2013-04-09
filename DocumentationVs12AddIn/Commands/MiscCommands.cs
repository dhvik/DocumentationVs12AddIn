using EnvDTE;

namespace DocumentationVs12AddIn.Commands {
	/// <summary>
	/// Summary description for MiscCommands.
	/// </summary>
	/// <remarks>
	/// 2013-04-08 dan: Created
	/// </remarks>
	public class MiscCommands:CommandBase {
		[Command("Text Editor::Ctrl+Shift+k")]
		public void KillLine() {
			if (Selection == null)
				return;
			Selection.StartOfLine();
			Selection.EndOfLine(true);
			Selection.Delete();
		}
	}
}