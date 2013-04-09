using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text.RegularExpressions;
using EnvDTE;
using EnvDTE80;
using log4net;

namespace DocumentationVs12AddIn.Commands {
	/// <summary>
	/// A base class for Command classes
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
		#region protected TextSelection Selection
		/// <summary>
		/// Gets the Selection of the CommandBase
		/// </summary>
		/// <value></value>
		protected TextSelection Selection {
			get { return ((TextSelection)DTE.ActiveDocument.Selection); }
		}
		#endregion
		#region protected bool IsHtml
		/// <summary>
		/// Gets the IsHtml of the CommandBase
		/// </summary>
		/// <value></value>
		protected bool IsHtml {
			get { return DTE.ActiveWindow.Document.Language == "HTML"; }
		}
		#endregion
		#region protected bool IsBasic
		/// <summary>
		/// Gets the IsBasic of the CommandBase
		/// </summary>
		/// <value></value>
		protected bool IsBasic {
			get { return DTE.ActiveWindow.Document.Language == "Basic"; }
		}
		#endregion
		#region protected bool IsCSharp
		/// <summary>
		/// Gets the IsCSharp of the CommandBase
		/// </summary>
		/// <value></value>
		protected bool IsCSharp {
			get { return DTE.ActiveWindow.Document.Language == "CSharp"; }
		}
		#endregion
		#region protected ILog Log
		/// <summary>
		/// Get/Sets the Log of the CommandBase
		/// </summary>
		/// <value></value>
		protected ILog Log { get; private set; }
		#endregion
		/* *******************************************************************
		 *  Constructors
		 * *******************************************************************/
		#region protected CommandBase()
		/// <summary>
		/// Initializes a new instance of the <b>CommandBase</b> class.
		/// </summary>
		protected CommandBase() {
			Log = LogManager.GetLogger(GetType());
		}
		#endregion
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
		#region protected Point GetPoint(TextPoint tp = null)
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
		#endregion
		#region protected void MoveToPoint(Point p, bool extend = false)
		/// <summary>
		///  Moves the current Selection to the supplied point (offset, line) (x,y)
		/// </summary>
		/// <param name="p"></param>
		/// <param name="extend"></param>
		protected void MoveToPoint(Point p, bool extend = false) {
			Selection.MoveToLineAndOffset(p.Y, p.X, extend);
		}
		#endregion
		#region protected string CutLeftWord()
		/// <summary>
		/// 
		/// </summary>
		/// <returns></returns>
		protected string CutLeftWord() {

			//get templateName
			Selection.WordLeft(true);
			string word = Selection.Text;

			//delete templateName
			Selection.Delete();

			return word;
		}
		#endregion
		#region protected string GetSeeXmlDocTag(string type)
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
		#endregion
		#region protected TextSelection FixSelection()
		/// <summary>
		/// Fixes the selection so that the selection starts at the first column of the first line
		/// and the active point is at the bottom. if we have started to mark part of the last line, inlcude the whole line
		/// </summary>
		/// <returns></returns>
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
		#endregion
		#region protected bool IsBetween(Point p, Point startPoint, Point endPoint)
		/// <summary>
		/// Checks if a point is between the given start and end points
		/// </summary>
		/// <param name="p">The <see cref="Point"/> to check</param>
		/// <param name="startPoint">The Startpoint</param>
		/// <param name="endPoint">The endpoint</param>
		/// <returns>True if the point is between the start and end points, false otherwize.</returns>
		protected bool IsBetween(Point p, Point startPoint, Point endPoint) {
			//if point is not in range
			if (p.Y < startPoint.Y | p.Y > endPoint.Y) {
				return false;
			}

			// all points on the same line
			if (startPoint.Y == endPoint.Y & p.Y == startPoint.Y) {
				if (startPoint.X < endPoint.X & p.X >= startPoint.X & p.X < endPoint.X) {
					return true;
				} else {
					return false;
				}
			}

			// start point and endpoint on different lines

			if (startPoint.Y < endPoint.Y) {
				//if point and enpoint on the same line
				if (p.Y == endPoint.Y) {
					//make sure that point is before endpoint
					if (p.X < endPoint.X) {
						return true;
					} else {
						return false;
					}
				}
				//if point and startpoint is on the same line
				if (p.Y == startPoint.Y) {
					//make sure that point is after startpoint
					if (p.X >= startPoint.X) {
						return true;
					} else {
						return false;
					}
				}
				//else it is between
				return true;
			}
			//else 
			return false;

		}
		#endregion
		#region protected CodeElement GetElement()
		/// <summary>
		/// Gets the CodeElement where the cursor is positioned.
		/// </summary>
		/// <returns>The found element</returns>
		protected CodeElement GetElement() {
			Point p = GetPoint();
			TextSelection sel = (TextSelection)DTE.ActiveDocument.Selection;
			object o = null;

			do {
				o = sel.ActivePoint.CodeElement[vsCMElement.vsCMElementProperty];
				if (((o != null))) {
					break; // TODO: might not be correct. Was : Exit Do
				}
				o = sel.ActivePoint.CodeElement[vsCMElement.vsCMElementFunction];
				if (((o != null))) {
					break; // TODO: might not be correct. Was : Exit Do
				}
				o = sel.ActivePoint.CodeElement[vsCMElement.vsCMElementVariable];
				if (((o != null))) {
					break; // TODO: might not be correct. Was : Exit Do
				}
				o = sel.ActivePoint.CodeElement[vsCMElement.vsCMElementEnum];
				if (((o != null))) {
					break; // TODO: might not be correct. Was : Exit Do
				}
				o = sel.ActivePoint.CodeElement[vsCMElement.vsCMElementEvent];
				if (((o != null))) {
					break; // TODO: might not be correct. Was : Exit Do
				}
				o = sel.ActivePoint.CodeElement[vsCMElement.vsCMElementStruct];
				if (((o != null))) {
					break; // TODO: might not be correct. Was : Exit Do
				}
				o = sel.ActivePoint.CodeElement[vsCMElement.vsCMElementDelegate];
				if (((o != null))) {
					break; // TODO: might not be correct. Was : Exit Do
				}
				o = sel.ActivePoint.CodeElement[vsCMElement.vsCMElementInterface];
				if (((o != null))) {
					break; // TODO: might not be correct. Was : Exit Do
				}
				o = sel.ActivePoint.CodeElement[vsCMElement.vsCMElementClass];
				if (((o != null))) {
					break; // TODO: might not be correct. Was : Exit Do
				}
				o = sel.ActivePoint.CodeElement[vsCMElement.vsCMElementNamespace];
				if (((o != null))) {
					break; // TODO: might not be correct. Was : Exit Do
				}
				break; // TODO: might not be correct. Was : Exit Do
			} while (true);
			if (o == null) {
				return null;
			}
			return (CodeElement)o;
		}
		#endregion
		#region protected void PasteSeeXmlDocParamTag()
		/// <summary>
		/// Pastes a See xml doc tag in the current xml documentation node
		/// </summary>
		protected void PasteSeeXmlDocParamTag() {
			Point p = GetPoint();
			TextSelection sel = Selection;

			dynamic xmlDocTag = GetCurrentXmlDocTag();
			if (xmlDocTag == null) {
				return;
			}
			CodeTypeRef type = default(CodeTypeRef);

			CodeElement ce = GetNextElement();

			string s = GetTagName(xmlDocTag);
			switch (s) {
				case "param":
					//get name of param
					Match m = Regex.Match(xmlDocTag, "name=\"(.+?)\"");

					if (!m.Success) {
						return;
					}
					string paramName = m.Groups[1].Value;
					type = GetParameterType(ce, paramName);

					break;
				case "summary":
				case "value":
				case "returns":
					type = GetReturnValueType(ce);

					break;

			}

			if ((type != null)) {
				sel.Insert(GetSeeXmlDocTagByType(type) + " ");
			}


		}
		#endregion
		#region protected string GetSeeXmlDocTagByType(CodeTypeRef type)
		/// <summary>
		/// Gets a xml documentation see tag by supplying a CodeTypeRef
		/// </summary>
		/// <param name="type">The type to get the see tag of</param>
		/// <returns>The see tag</returns>
		protected string GetSeeXmlDocTagByType(CodeTypeRef type) {
			return GetSeeXmlDocTag(GetTypeName(type));
		}
		#endregion
		#region protected string GetTypeName(CodeTypeRef type)
		/// <summary>
		/// Gets the name of the supplied type.
		/// </summary>
		/// <param name="type">The type</param>
		/// <returns>The name of the type</returns>
		protected string GetTypeName(CodeTypeRef type) {
			string name = null;
			try {
				name = type.CodeType.FullName;
			} catch (Exception ex) {
				name = type.AsString;
			}
			string finalName = null;
			char[] separatorCharacters = new char[] {
				'<',
				'>',
				','
			};
			List<string> importedNamespaces = GetImportedNamespaces();
			while ((name.Length > 0)) {
				int i = name.IndexOfAny(separatorCharacters);
				if (i == -1) {
					finalName += RemoveRedundantQualifiersFromTypeName(name, importedNamespaces);
					name = "";
				} else if (i == 0) {
					finalName += name.Substring(0, 1);
					name = name.Substring(1);
				} else {
					finalName += RemoveRedundantQualifiersFromTypeName(name.Substring(0, i), importedNamespaces) + name.Substring(i, 1);
					name = name.Substring(i + 1);
				}
			}
			return finalName;
		}
		#endregion
		#region private string RemoveRedundantQualifiersFromTypeName(string typeName, List<string> importedNamespaces)
		/// <summary>
		/// Shortens the type name to minimize the type name
		/// </summary>
		/// <param name="typeName">The name of the type</param>
		/// <param name="importedNamespaces">The <see cref="List{t}"/> of imported namespaces</param>
		/// <returns></returns>
		private string RemoveRedundantQualifiersFromTypeName(string typeName, List<string> importedNamespaces) {
			if (typeName == null) {
				return null;
			}
			string finalName = "";
			typeName = typeName.Trim();

			while ((typeName.Length > 0)) {
				dynamic index = typeName.LastIndexOf(".");
				if (index == -1) {
					finalName = typeName + finalName;
					break; // TODO: might not be correct. Was : Exit While
				} else {
					finalName = typeName.Substring(index) + finalName;
					typeName = typeName.Substring(0, index);
					if (importedNamespaces.Contains(typeName)) {
						break; // TODO: might not be correct. Was : Exit While
					}
				}
			}
			if ((finalName.StartsWith("."))) {
				finalName = finalName.Substring(1);
			}
			return finalName;
		}
		#endregion
		#region private List<string> GetImportedNamespaces()
		/// <summary>
		/// Gets the imported namespaces for the current document.
		/// </summary>
		/// <returns></returns>
		private List<string> GetImportedNamespaces() {
			List<string> list = new List<string>();
			DTE2 dte2 = (DTE2)DTE;
			FileCodeModel2 fileCM = (FileCodeModel2)dte2.ActiveDocument.ProjectItem.FileCodeModel;
			CodeElement2 elt = default(CodeElement2);
			int i = 0;
			//MsgBox("About to walk top-level elements ...")
			for (i = 1; i <= fileCM.CodeElements.Count; i++) {
				elt = (CodeElement2)fileCM.CodeElements.Item(i);
				if (elt.Kind == vsCMElement.vsCMElementImportStmt) {
					CodeImport imp = (CodeImport)elt;
					list.Add(imp.Namespace);
				} else if (elt.Kind == vsCMElement.vsCMElementNamespace) {
					CodeNamespace ns = (CodeNamespace)elt;
					list.Add(ns.FullName);
				}
			}
			return list;
		}
		#endregion
		#region private string GetCurrentXmlDocTag()
		/// <summary>
		/// Gets the current XmlDoc tag
		/// </summary>
		/// <returns>The complete xml tag that we are in</returns>
		private string GetCurrentXmlDocTag() {
			Point p = GetPoint();
			TextSelection sel = (TextSelection)DTE.ActiveWindow.Document.Selection;

			Point startPoint = Point.Empty;
			Point endPoint = Point.Empty;

			string tag = null;
			//if (sel.FindPattern("\\<[^\\>]+\\>", vsFindOptions.vsFindOptionsBackwards + vsFindOptions.vsFindOptionsRegularExpression)) {
			if (sel.FindPattern(@"\<[^\>]+\>", (int)(vsFindOptions.vsFindOptionsBackwards + (int)vsFindOptions.vsFindOptionsRegularExpression))) {
				startPoint = GetPoint();
				tag = sel.Text;

				string tagName = GetTagName(tag);
				if (tagName == null) {
					return null;
				}

				if (sel.FindPattern("\\</" + tagName + "\\>", (int)vsFindOptions.vsFindOptionsRegularExpression)) {
					endPoint = GetPoint();
				}
			}
			if (!IsBetween(p, startPoint, endPoint)) {
				tag = null;
			}
			MoveToPoint(p);
			return tag;
		}
		#endregion
		#region private string GetTagName(string tag)
		/// <summary>
		/// Gets the tagname of the supplied xml tag
		/// </summary>
		/// <param name="tag">The xml tag</param>
		/// <returns>The element name of the tag</returns>
		private string GetTagName(string tag) {
			Match m = Regex.Match(tag, "<([^>\\s]+)");
			if (m.Success) {
				return m.Groups[1].Value;
			}
			return null;
		}
		#endregion
		#region private CodeTypeRef GetParameterType(CodeElement ce, string parameterName)
		/// <summary>
		/// Gets the type of the parameter
		/// </summary>
		/// <param name="ce">The codeElement (CodeFunction) that has the parameter</param>
		/// <param name="parameterName">The name of the parameter</param>
		/// <returns>The type that the parameter has, or nothing if the parameter was not found</returns>
		private CodeTypeRef GetParameterType(CodeElement ce, string parameterName) {
			switch (ce.Kind) {
				case vsCMElement.vsCMElementFunction:
					CodeFunction cf = (CodeFunction)ce;
					foreach (CodeParameter p in cf.Parameters) {
						if (p.Name == parameterName) {
							return p.Type;
						}
					}

					break;
				case vsCMElement.vsCMElementDelegate:
					CodeDelegate cd = (CodeDelegate)ce;
					foreach (CodeParameter p in cd.Parameters) {
						if (p.Name == parameterName) {
							return p.Type;
						}
					}

					break;
			}
			return null;
		}
		#endregion
		#region private CodeTypeRef GetReturnValueType(CodeElement ce)
		/// <summary>
		/// Gets the returnvalue type of the given CodeElement
		/// </summary>
		/// <param name="ce">The code element to retrieve the return type from</param>
		/// <returns>The <see cref="CodeTypeRef"/> that is the returnvalue of the element</returns>
		private CodeTypeRef GetReturnValueType(CodeElement ce) {
			switch (ce.Kind) {
				case vsCMElement.vsCMElementFunction:
					return ((CodeFunction)ce).Type;
				case vsCMElement.vsCMElementProperty:
					return ((CodeProperty)ce).Type;
				case vsCMElement.vsCMElementVariable:
					return ((CodeVariable)ce).Type;
			}
			return null;
		}
		#endregion
		#region private CodeElement GetNextElement()
		/// <summary>
		/// Gets the next CodeElement that follows the cursor
		/// </summary>
		/// <returns>The found <see cref="CodeElement"/> or Nothing if no more elements could be found</returns>
		private CodeElement GetNextElement() {
			Point p = GetPoint();
			CodeElement ce = GetElement();

			//if we didn't get any element at the given position,
			//try to move to the next line and try again.
			if (ce == null) {
				Point p2 = new Point(1, p.Y);

				do {
					p2.Y = p2.Y + 1;
					MoveToPoint(p2);
					Point p3 = GetPoint();
					if (p2.Y != p3.Y) {
						break; // TODO: might not be correct. Was : Exit Do
					}
					ce = GetElement();
					if ((ce != null)) {
						if (ce.Kind == vsCMElement.vsCMElementParameter) {
							//ignore parameters
						} else {
							break; // TODO: might not be correct. Was : Exit Do
						}
					}
				} while (true);

				MoveToPoint(p);

				if (ce == null) {
					return null;
				} else {
					return ce;
				}
			}
			if (ce.Kind == vsCMElement.vsCMElementClass) {
				CodeClass cc = (CodeClass)ce;
				foreach (CodeElement elem in cc.Members) {
					if (elem.StartPoint.Line > p.Y) {
						return elem;
					}
				}
			} else if (ce.Kind == vsCMElement.vsCMElementInterface) {
				CodeInterface cc = (CodeInterface)ce;
				foreach (CodeElement elem in cc.Members) {
					if (elem.StartPoint.Line > p.Y) {
						return elem;
					}
				}
			} else if (ce.Kind == vsCMElement.vsCMElementStruct) {
				CodeStruct cc = (CodeStruct)ce;
				foreach (CodeElement elem in cc.Members) {
					if (elem.StartPoint.Line > p.Y) {
						return elem;
					}
				}
			} else if (ce.Kind == vsCMElement.vsCMElementNamespace) {
				CodeNamespace cn = (CodeNamespace)ce;
				foreach (CodeElement elem in cn.Members) {
					if (elem.StartPoint.Line > p.Y) {
						return elem;
					}
				}
			}
			return null;
		}
		#endregion
	}
}