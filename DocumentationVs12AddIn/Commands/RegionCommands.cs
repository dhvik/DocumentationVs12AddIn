using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using EnvDTE;

namespace DocumentationVs12AddIn.Commands {
	/// <summary>
	/// Summary description for RegionCommands.
	/// </summary>
	/// <remarks>
	/// 2013-04-08 dan: Created
	/// </remarks>
	public class RegionCommands : CommandBase {

		private bool _lastCollapsedAllRegionsNotCompleted;


		[Command("Text Editor::Ctrl+Shift++")] 
		public void ExpandAllRegions() {
			if (IsHtml) {
				DTE.ExecuteCommand("Edit.StopOutlining");
				OutlineRegions();
			} else {
				DTE.ExecuteCommand("Edit.StopOutlining");
				DTE.ExecuteCommand("Edit.StartAutomaticOutlining");
			}

		}

		private void OutlineRegions() {
			//var selection = (TextSelection)DTE.ActiveDocument.Selection;

			const string REGION_START = "//#region";
			const string REGION_END = "//#endregion";

			Selection.SelectAll();
			var text = Selection.Text;
			Selection.StartOfDocument(true);

			var lastIndex = 0;
			var startRegions = new Stack<int>();

			do {
				var startIndex = text.IndexOf(REGION_START, lastIndex);
				var endIndex = text.IndexOf(REGION_END, lastIndex);

				if (startIndex == -1 && endIndex == -1) {
					break;
				}

				if (startIndex != -1 && startIndex < endIndex) {
					startRegions.Push(startIndex);
					lastIndex = startIndex + 1;
				} else {
					//Outline region...
					Selection.MoveToLineAndOffset(CalcLineNumber(text, startRegions.Pop()), 1);
					Selection.MoveToLineAndOffset(CalcLineNumber(text, endIndex) + 1, 1, true);
					Selection.OutlineSection();
					DTE.ExecuteCommand("Edit.ToggleOutliningExpansion");

					lastIndex = endIndex + 1;
				}
			} while (true);

			Selection.StartOfDocument();

		}

		private int CalcLineNumber(string text, int index) {
			var lineNumber = 1;
			var i = 0;

			while (i < index) {
				if (text[i] == '\n') {
					lineNumber++;
					//i++;
				}
				i++;
			}
			return lineNumber;
		}



		[Command("Text Editor::Ctrl+Shift+-")]
		public void CollapseAllRegions() {
			//find point to go to when operation is done.
			var p = GetPoint();
			var sel = Selection;
			//move to end of line incase we are standing in the middle of a region directive.
			//Fix supplied by DragonWang_SFIS
			sel.EndOfLine();
			//Try
			try {
				//	'DTE.SuppressUI = True

				if (!_lastCollapsedAllRegionsNotCompleted) {
					//Loop until you have exited all nested region blocks. 
					bool done;
					do {
						done = true;
						//if inside region, move to start of region
						if (!sel.FindPattern("#region", (int) vsFindOptions.vsFindOptionsBackwards)) 
							continue;
						sel.StartOfLine();
						sel.Collapse();
						var p2 = GetPoint();
						if (!sel.FindPattern("#endregion")) 
							continue;
						var p3 = GetPoint();
						//are we inside a region block?
						if (p2.Y > p3.Y) 
							continue;
						p = p2;
						//Move to start of region
						MoveToPoint(p2);
						//and redo the search
						done = false;
					} while (!done);

					ExpandAllRegions();
					//To collapse nested regions, we need to iterate region blocks from the end of the
					//file and move upwards.
					//Fix supplied by DragonWang_SFIS
					sel.EndOfDocument();
					_lastCollapsedAllRegionsNotCompleted = true;
				}
				//2006-05-30 Fixed ignoring commented regions.
				//while (sel.FindPattern("^:b*/*\\#region",(int)(vsFindOptions.vsFindOptionsBackwards + (int) vsFindOptions.vsFindOptionsRegularExpression))) {
				//vs 2012 uses normal regexp expressions
				while (sel.FindPattern(@"^\s*/*\#region",(int)(vsFindOptions.vsFindOptionsBackwards + (int) vsFindOptions.vsFindOptionsRegularExpression))) {
					sel.StartOfLine();
					DTE.ExecuteCommand("Edit.ToggleOutliningExpansion");
					//		'Try
					//		'    '2006-10-26 , Dan Händevik; if region is a commented one /* */ then we would collapse the closest overlying region.
					//		'    'make sure that we are on the start line of the collapsed line...	
					//		'    '   sel.StartOfLine(vsStartOfLineOptions.vsStartOfLineOptionsFirstColumn)
					//		'Catch ex As Exception
					//		'    System.Threading.Thread.Sleep(1000)
					//		'    sel = DTE.ActiveDocument.Selection
					//		'    '   sel.StartOfLine(vsStartOfLineOptions.vsStartOfLineOptionsFirstColumn)
					//		'End Try

				}
				//restore old point
				MoveToPoint(p);
				_lastCollapsedAllRegionsNotCompleted = false;
			} catch (Exception) {
				MessageBox.Show(
					string.Format("Visual Studio threw an exception so CollapseAllRegions couldn't complete.{0}Please run this macro again to continue the function.", Environment.NewLine), "Macro couldn't complete, try again!");
				//MessageBox.Show(objE.Message + Environment.NewLine + objE.StackTrace, "Failed to collapse region")
			}
		}

		[Command("Text Editor::Ctrl+m,Ctrl+m")]
		public void ToggleParentRegion() {
			//Dim objSelection As TextSelection
			var objSelection = (TextSelection)DTE.ActiveDocument.Selection;
			var point = GetPoint();
			objSelection.EndOfLine();
			var endLine = objSelection.ActivePoint.Line;
			objSelection.LineUp();

			//is the selected line a collapsed outline region?
			if (endLine - 1 != objSelection.ActivePoint.Line) {
				MoveToPoint(point);
				DTE.ExecuteCommand("Edit.ToggleOutliningExpansion");
			} else {
				MoveToPoint(point);
				objSelection.SelectLine();
				if (objSelection.Text.ToLower().Contains("#region")) {
					//Updated 14.08.2003
					MoveToPoint(point);
					DTE.ExecuteCommand("Edit.ToggleOutliningExpansion");
					Selection.StartOfLine();

				} else if (objSelection.FindText("#region", (int)vsFindOptions.vsFindOptionsBackwards)) {

					Selection.StartOfLine();
					DTE.ExecuteCommand("Edit.ToggleOutliningExpansion");
					Selection.StartOfLine();
				} else {
					MoveToPoint(point);
				}

			}
		}
	}
}