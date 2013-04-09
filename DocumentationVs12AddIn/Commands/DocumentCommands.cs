using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;
using EnvDTE;
using EnvDTE80;

namespace DocumentationVs12AddIn.Commands {
	/// <summary>
	/// Summary description for DocumentCommands.
	/// </summary>
	/// <remarks>
	/// 2013-04-09 dan: Created
	/// </remarks>
	public class DocumentCommands : CommandBase {
		[Command("Text Editor::Ctrl+Alt+d")]
		public void DocumentThis() {
			DocumentAndRegionize(false);
		}
		[Command("Text Editor::Ctrl+d")]
		public void DocumentAndRegionizeThis() {
			DocumentAndRegionize(true);
		}

		protected CodeElement DocumentAndRegionize(bool regionize) {
			if (IsHtml) {
				DocumentAndRegionizeHtml(regionize);
				return null;
			}
			var ce = GetDocumentationElement();
			if (ce == null)
				return null;

			var startPoint = GetPoint();
			try {
				var xdoc = PrepareXMLDocumentation(ce);
				if (xdoc == null) {
					MoveToPoint(startPoint);
					return null;
				}
				var summaryNode = AutoDocumentItem(ce, xdoc);
				SaveXMLDocumentation(xdoc, ce);
				SetCursorToNodeText(summaryNode);
				if (regionize && ShouldCodeElementBeRegionized(ce)) {
					RegionizeCodeElement(ce);
				}
			} catch (Exception e) {
				MessageBox.Show(e.Message + Environment.NewLine + e.StackTrace, "Documentator", MessageBoxButtons.OK,
								MessageBoxIcon.Hand);
				//if anything goes wrong, return to old position
				MoveToPoint(startPoint);
				ce = null;
			}
			return ce;
		}

		private void RegionizeCodeElement(CodeElement elem) {

			string declaration = GetDeclaration(elem);

			TextSelection sel = (TextSelection)DTE.ActiveWindow.Selection;

			//store startpoint as reference
			Point p = GetPoint(sel.ActivePoint);


			//search for xmldocumentation
			Point ep = GetPoint(elem.EndPoint);
			sel.MoveToPoint(elem.StartPoint);
			Point summaryPoint = Point.Empty;
			Point inheritdocPoint = Point.Empty;
			if (sel.FindPattern("<summary>", (int)vsFindOptions.vsFindOptionsBackwards)) {
				sel.StartOfLine();
				summaryPoint = GetPoint(sel.ActivePoint);
				MoveToPoint(p);
			}
			if (sel.FindPattern("<inheritdoc", (int)vsFindOptions.vsFindOptionsBackwards)) {
				sel.StartOfLine();
				inheritdocPoint = GetPoint(sel.ActivePoint);
				MoveToPoint(p);
			}
			Point regionPoint = summaryPoint;
			if (inheritdocPoint != Point.Empty) {
				if (regionPoint == Point.Empty | regionPoint.Y < inheritdocPoint.Y) {
					regionPoint = inheritdocPoint;
				}
			}

			if (regionPoint != Point.Empty) {
				MoveToPoint(regionPoint);

				sel.StartOfLine();

				// do we have a region specified already, then remove all whitespaces between it and the xml comment
				//if (sel.FindPattern("^:b*\\#region:Wh", (int) (vsFindOptions.vsFindOptionsRegularExpression + (int) vsFindOptions.vsFindOptionsBackwards))) {
				if (sel.FindPattern(@"^\s*\#region\s", (int)(vsFindOptions.vsFindOptionsRegularExpression + (int)vsFindOptions.vsFindOptionsBackwards))) {
					sel.EndOfLine();
					Point prevRegionPoint = GetPoint(sel.ActivePoint);
					MoveToPoint(regionPoint, true);
					//if this region is connected to this item (only whitespaces 
					if (Regex.IsMatch(sel.Text, "^[\\12\\15\\s\\t]*$", RegexOptions.Singleline)) {
						sel.Delete();
						sel.Insert(Environment.NewLine);
						//update the regionPoint
						MoveToPoint(prevRegionPoint);
						sel.StartOfLine();
						sel.LineDown();
						regionPoint = GetPoint(sel.ActivePoint);
					} else {
						//wrap region again, since we wrapped it up by searching...
						MoveToPoint(prevRegionPoint);

						DTE.ExecuteCommand("Edit.ToggleOutliningExpansion");
						MoveToPoint(regionPoint);
					}
				} else {
					MoveToPoint(regionPoint);
				}


				//select the line above the documentation (or start of element)
				sel.LineUp();
				sel.SelectLine();

				//first check that the prev line is not several lines (folded)
				//is the region missing?

				if (sel.AnchorPoint.Line != regionPoint.Y - 1 | sel.Text.ToLower().IndexOf("#region", StringComparison.Ordinal) == -1) {
					MoveToPoint(regionPoint);
					if (IsBasic) {
						sel.Insert("#Region \"" + declaration + "\"" + Environment.NewLine);
					} else {
						sel.Insert("\t\t#region " + declaration + Environment.NewLine);
					}

					//add linecount if new line is added above start position
					if (sel.ActivePoint.Line <= p.Y) {
						p.Y += 1;
					}
					if (sel.ActivePoint.Line <= ep.Y) {
						ep.Y += 1;
					}

					//move to endregion position
					MoveToPoint(ep);
					//sel.MoveToPoint(elem.EndPoint)

					//Check for lines with underscore variable decorations
					//Will include all underscore variables trailing the regionized declaration
					bool searchForMoreVariables = true;
					do {
						Point tmpPoint = GetPoint();
						sel.LineDown();
						sel.StartOfLine();
						sel.SelectLine();
						if (sel.Text.IndexOf("_") != -1) {
							sel.SwapAnchor();
							sel.Collapse();
							sel.StartOfLine();
							sel.FindPattern(":b_", (int)vsFindOptions.vsFindOptionsRegularExpression);
							CodeElement ce = sel.ActivePoint.CodeElement[vsCMElement.vsCMElementVariable];
							if ((ce == null)) {
								MoveToPoint(tmpPoint);
								searchForMoreVariables = false;

							} else if ((ce.Kind != vsCMElement.vsCMElementVariable)) {
								MoveToPoint(tmpPoint);
								searchForMoreVariables = false;
							}
						} else {
							MoveToPoint(tmpPoint);
							searchForMoreVariables = false;
						}
					} while (searchForMoreVariables);

					sel.EndOfLine();
					sel.Collapse();

					//add endregion
					if (IsBasic) {
						sel.Insert(Environment.NewLine + "#End Region");
					} else {
						sel.Insert(Environment.NewLine + "\t\t#endregion");
					}


					//add linecount if new line is added above start position
					if (sel.ActivePoint.Line <= p.Y) {
						p.Y += 1;
					}
				} else {
					//region exists, update it.
					//sel.Collapse()
					sel.SwapAnchor();
					sel.StartOfLine();
					regionPoint = GetPoint();
					sel.SelectLine();
					sel.Delete();

					if (IsBasic) {
						sel.Insert("#Region \"" + declaration + "\"" + Environment.NewLine);
					} else {
						sel.Insert("\t" + "\t" + "#region " + declaration + Environment.NewLine);
					}

					//put marker at the end of the region
					if (IsBasic) {
						sel.FindPattern("#End Region");
					} else {
						sel.FindPattern("#endregion");
					}

					sel.Collapse();
					sel.StartOfLine();
					Point endRegionPoint = GetPoint(sel.ActivePoint);

					//remove any leading whitespaces (blank lines)

					//if (sel.FindPattern("[^:b:Wh]", vsFindOptions.vsFindOptionsRegularExpression + vsFindOptions.vsFindOptionsBackwards + vsFindOptions.vsFindOptionsMatchInHiddenText)) {
					if (sel.FindPattern(@"[^\s]", (int)vsFindOptions.vsFindOptionsRegularExpression + (int)vsFindOptions.vsFindOptionsBackwards + (int)vsFindOptions.vsFindOptionsMatchInHiddenText)) {
						sel.EndOfLine();
						MoveToPoint(endRegionPoint, true);
						sel.Delete();
						sel.Insert(Environment.NewLine);
						sel.EndOfLine();
					}

				}
				//select whole region
				MoveToPoint(regionPoint, true);

				//format the selection
				DTE.ExecuteCommand("Edit.FormatSelection");
			}

			//go back to old point
			MoveToPoint(p);
		}

		private string GetDeclaration(CodeElement elem) {
			TextSelection sel = (TextSelection)DTE.ActiveWindow.Selection;

			Point p = GetPoint(sel.ActivePoint);

			sel.MoveToPoint(elem.StartPoint);
			sel.MoveToPoint(elem.EndPoint, true);

			string str = sel.Text;
			MoveToPoint(p);
			string rexnl = Regex.Escape(Environment.NewLine);
			//2009-02-24, Dan Händevik. Now supports comments in declaration 
			string rexAttribute = "(?:\\[[^\\]]*]\\s*)";
			string rexOneLineComment = "(?:\\/{2,}[^" + rexnl + "]*" + rexnl + "\\s*)";
			string rexMultiLineComment = "(?:\\/\\*.*?\\*\\/\\s*)";
			string rexAttributeOrComment = "(?:" + rexAttribute + "|" + rexOneLineComment + "|" + rexMultiLineComment + ")";
			if (IsBasic) {
				if (elem.Kind == vsCMElement.vsCMElementVariable) {
					str = Regex.Replace(str, "^\\s*(?:(?:\\<:=[^\\>]*>\\s*)*)?([^\\=" + rexnl + "]*)\\s*\\=?.*;.*$", "$1", RegexOptions.Singleline);
				} else {
					str = Regex.Replace(str, "^\\s*(?:(?:\\<:=[^\\>]*>\\s*)*)?([^" + rexnl + "]*).*$", "$1", RegexOptions.Singleline);
				}
			} else {
				if (elem.Kind == vsCMElement.vsCMElementVariable) {
					str = Regex.Replace(str, "^\\s*(?:" + rexAttributeOrComment + "*)?([^\\=" + rexnl + "]*)\\s*\\=?.*;.*$", "$1", RegexOptions.Singleline);
				} else {
					str = Regex.Replace(str, "^\\s*(?:" + rexAttributeOrComment + "*)?([^\\{:;]*)\\s*.*$", "$1", RegexOptions.Singleline);

				}
			}
			//strip tabs/newlines from declaration 
			str = Regex.Replace(str, "[\\012\\015\\t]*", "", RegexOptions.Singleline);
			//remove double whitespaces
			str = Regex.Replace(str, "\\s+", " ", RegexOptions.Singleline);
			//remove where clauses...
			str = Regex.Replace(str, "where\\s+\\w+\\s+$", "", RegexOptions.Singleline);
			//trim declaration
			str = str.Trim();

			return str;
		}

		private bool ShouldCodeElementBeRegionized(CodeElement element) {
			Type type = GetElementType(element);
			if (object.ReferenceEquals(type, typeof(CodeVariable))) {
				CodeVariable cv = (CodeVariable)element;
				CodeElement parent = (CodeElement)cv.Parent;
				//variables should not be regionized of parent isn't a class or a struct
				if (parent.Kind != vsCMElement.vsCMElementClass & parent.Kind != vsCMElement.vsCMElementStruct) {
					return false;
				}

			} else if (object.ReferenceEquals(type, typeof(CodeEvent))) {
				CodeEvent ce = (CodeEvent)element;
				CodeElement parent = (CodeElement)ce.Parent;
				//events should not be regionized of parent isn't a class or a struct
				if (parent.Kind != vsCMElement.vsCMElementClass & parent.Kind != vsCMElement.vsCMElementStruct & parent.Kind != vsCMElement.vsCMElementInterface) {
					return false;
				}

			} else if (object.ReferenceEquals(type, typeof(CodeInterface))) {
				CodeInterface ci = (CodeInterface)element;
				CodeElement parent = (CodeElement)ci.Parent;
				//interfaces should not be regionized of parent isn't a class or a struct
				if (parent.Kind != vsCMElement.vsCMElementClass & parent.Kind != vsCMElement.vsCMElementStruct) {
					return false;
				}
			} else if (object.ReferenceEquals(type, typeof(CodeClass))) {
				return false;
			}
			return true;
		}

		protected Type GetElementType(CodeElement element) {
			switch (element.Kind) {
				case vsCMElement.vsCMElementFunction:
					return typeof(CodeFunction);
				case vsCMElement.vsCMElementProperty:
					return typeof(CodeProperty);
				case vsCMElement.vsCMElementVariable:
					return typeof(CodeVariable);
				case vsCMElement.vsCMElementEnum:
					return typeof(CodeEnum);
				case vsCMElement.vsCMElementEvent:
					return typeof(CodeEvent);
				case vsCMElement.vsCMElementDelegate:
					return typeof(CodeDelegate);
				case vsCMElement.vsCMElementClass:
					return typeof(CodeClass);
				case vsCMElement.vsCMElementStruct:
					return typeof(CodeStruct);
				case vsCMElement.vsCMElementInterface:
					return typeof(CodeInterface);
				default:
					return null;
				//throw new ApplicationException("Unknown element kind "+element.Kind);
			}
		}

		#region "Private Function SetCursorToNodeText(ByVal node As XmlNode)"
		/// <summary>
		///	 Sets the position of the cursor to the supplied node of the xml
		/// documentation
		/// </summary>
		/// <param name="node" ></param>
		/// <returns ></returns>
		protected void SetCursorToNodeText(XmlNode node, bool positionLast = false, bool searchForCursorMark = false) {
			var sel = (TextSelection)DTE.ActiveWindow.Selection;
			bool emptyNode = node.InnerText.Length == 0 | Regex.IsMatch(node.InnerText, "^\\s*$");

			//position the cursor at the correct position
			sel.FindPattern("<" + node.Name + ">", (int)vsFindOptions.vsFindOptionsBackwards);
			if (emptyNode) {
				sel.Collapse();
				Point p = GetPoint();
				sel.SelectLine();
				string text = sel.Text;
				MoveToPoint(p);

				Match m;
				m = Regex.Match(text, IsBasic ? "^(\\s*''')" : "^(\\s*///)");
				if (m.Success) {
					dynamic nrOfLines = CountLineBreaks(node.InnerText);
					if (nrOfLines < 3) {
						sel.Insert(Environment.NewLine + m.Groups[1].Value + " ");

						//add an extra line if inner text is nothing.
						if (node.InnerText.Length == 0) {
							p = GetPoint();
							sel.Insert(Environment.NewLine + m.Groups[1].Value + " ");
							MoveToPoint(p);
						}
					}
				}
			} else {
				if (positionLast) {
					sel.Collapse();
					Point nodeStartPoint = GetPoint();
					sel.FindPattern("</" + node.Name + ">");
					sel.SwapAnchor();
					sel.Collapse();
					//If we should look for the | marker
					if (searchForCursorMark) {
						Point nodeEndPoint = GetPoint();
						sel.FindPattern("|", (int)vsFindOptions.vsFindOptionsBackwards);
						Point p = GetPoint();
						if (IsBetween(p, nodeStartPoint, nodeEndPoint)) {
							sel.Delete();
						} else {
							MoveToPoint(nodeEndPoint);
						}
					}
				} else {
					sel.SwapAnchor();
					sel.Collapse();
					sel.LineDown();
				}
			}

		}
		#endregion
		#region "Private Function IsBetween(ByVal p As Point, ByVal startPoint As Point, ByVal endPoint As Point) As Boolean"

		#endregion
		#region "Private Function CountLineBreaks(ByVal text As String) As Integer"
		/// <summary>
		/// Counts the number of linebreaks in the text
		/// </summary>
		/// <param name="text">The text to search</param>
		/// <returns>The number of linebreaks</returns>
		private int CountLineBreaks(string text) {
			if (text == null) {
				return 0;
			}
			int nr = 0;
			int p = -1;
			do {
				p = text.IndexOf(Environment.NewLine, p + 1, System.StringComparison.Ordinal);
				if (p == -1) {
					break; // TODO: might not be correct. Was : Exit Do
				}
				nr = nr + 1;
			} while (true);
			return nr;
		}
		#endregion
		#region "Private Function SaveXMLDocumentation(ByVal xdoc As XmlDocument, ByVal element As CodeElement)"
		/// <summary>
		/// Saves the supplied xml documentation to the element
		/// </summary>
		/// <param name="xdoc">The <see cref="XmlDocument"/> to save</param>
		/// <param name="element">The <see cref="CodeElement"/> to save it to.</param>
		/// <returns></returns>
		protected object SaveXMLDocumentation(XmlDocument xdoc, CodeElement element) {
			object functionReturnValue = null;
			if (xdoc == null) {
				return functionReturnValue;
			}
			System.Reflection.PropertyInfo pInfo = GetPropertyInfo(element, "DocComment");
			if (pInfo == null) {
				return null;
			}
			string xml = GetXmlDocumentation(xdoc);
			pInfo.SetValue(element, xml, null);
			return functionReturnValue;
		}
		#endregion
		#region "Private Function GetPropertyInfo(ByVal element As CodeElement, ByVal name As String) As Reflection.PropertyInfo"
		/// <summary>
		/// Gets the propertyinfo for the supplied method of the element.
		/// </summary>
		/// <param name="element">The <see cref="CodeElement"/> to get the property info from</param>
		/// <param name="name">The name of the property.</param>
		/// <returns>The found <see cref="Reflection.PropertyInfo"/> </returns>
		private PropertyInfo GetPropertyInfo(CodeElement element, string name) {
			Type type = GetElementType(element);
			if (type == null) {
				return null;
			}

			System.Reflection.PropertyInfo pInfo = type.GetProperty(name);
			//If pInfo Is Nothing Then
			//    Throw New ApplicationException("Cannot find property " + name + " in type " + type.FullName)
			//End If
			return pInfo;
		}
		#endregion
		#region "Private Function GetXmlDocumentation(ByVal xdoc As XmlDocument) As String"
		/// <summary>
		///	 Gets the xml from the xml document
		/// </summary>
		/// <param name="xdoc" ></param>
		/// <returns ></returns>
		private string GetXmlDocumentation(XmlDocument xdoc) {
			//split some singletags into start and endtags
			string xml = xdoc.InnerXml;
			xml = Regex.Replace(xml, "<(c|code|example|exception|list|listheader|term|description|param|permission|remarks|returns|summary|value|typeparam)\\b([^\\\\>]*?)\\s*/>", "<$1$2></$1>", RegexOptions.Singleline);
			if (IsBasic) {
				xml = FormatBasicXmlDocumentation(xml);
			}
			return xml;
		}
		#endregion
		#region "Private Function FormatBasicXmlDocumentation(ByVal xml As String) As String"
		/// <summary>
		/// Performs formatting of the supplied doc xml to lines starting with ''' 
		/// </summary>
		/// <param name="xml">The xml text to format</param>
		/// <returns>The formatted text</returns>
		private string FormatBasicXmlDocumentation(string xml) {
			xml = Regex.Replace(xml, "^<doc>(.*)</doc>$", "$1", RegexOptions.Singleline);
			xml = Regex.Replace(xml, "<summary>\\s*</summary>", "<summary>" + Environment.NewLine + "</summary>", RegexOptions.Multiline);
			xml = Regex.Replace(xml, "(</(?:summary|returns|param|exception|code|remarks)>)\\s*<", "$1" + Environment.NewLine + "<", RegexOptions.Multiline);
			//xml = Regex.Replace(xml, "^\s*", "''' ", RegexOptions.Multiline)
			return xml;
		}
		#endregion
		#region "Private Function AutoDocumentItem(ByVal elem As CodeElement, ByVal xdoc As XmlDocument, ByVal returnType As CodeTypeRef, ByVal parameters As CodeElements) As XmlNode"
		/// <summary>
		/// Creates documentation for the supplied element
		/// </summary>
		/// <param name="elem">The <see cref="CodeElement"/> to document</param>
		/// <param name="xdoc">The <see cref="XmlDocument"/> containing the documentation</param>
		/// <param name="returnType">The return <see cref="CodeTypeRef"/> of the element</param>
		/// <param name="parameters">The list of <see cref="CodeElements"/> sent as parameters</param>
		/// <returns>The last <see cref="XmlNode"/> of the documentation.</returns>
		private XmlNode AutoDocumentItem(CodeElement elem, XmlDocument xdoc) {
			RemoveEmptyElements(xdoc);

			//'Pay attention to the inheritdoc node. Don't autodocument anything if it is present
			if ((HasNode(xdoc, "/doc/inheritdoc"))) {
				return xdoc.DocumentElement.LastChild;
			}


			XmlNode summaryNode = EnsureNode(xdoc, "/doc/summary", null, true);
			XmlNode lastNode = summaryNode;
			XmlNode returnNode = default(XmlNode);
			XmlNode paramDocNode = default(XmlNode);
			CodeTypeRef returnType = default(CodeTypeRef);
			CodeElements parameters = default(CodeElements);
			System.Reflection.PropertyInfo parametersInfo = GetPropertyInfo(elem, "Parameters");
			if ((parametersInfo != null)) {
				parameters = (CodeElements)parametersInfo.GetValue(elem, null);
			}



			System.Reflection.PropertyInfo typeInfo = GetPropertyInfo(elem, "Type");
			if ((typeInfo != null)) {
				returnType = (CodeTypeRef)typeInfo.GetValue(elem, null);
			}
			//add parameter nodes
			if ((parameters != null)) {
				lastNode = AddParameters(xdoc, parameters, lastNode);
			}

			//Check for generic parameters
			Match m = Regex.Match(elem.FullName, "<([^>]*)>$");
			if (m.Success) {
				string[] typeMatches = Regex.Split(m.Groups[1].Value, "\\s*,\\s*");

				foreach (string name in typeMatches) {
					XmlNode node = EnsureNode(xdoc, "/doc/typeparam[@name='" + name + "']");
					if ((!object.ReferenceEquals(lastNode.NextSibling, node))) {
						SwapNodes(lastNode.NextSibling, node);
					}
					lastNode = node;
				}
			}

			switch (elem.Kind) {
				case vsCMElement.vsCMElementProperty:
					lastNode = AddXMLDocNode(xdoc, "value", lastNode);
					break; // TODO: might not be correct. Was : Exit Select
				case vsCMElement.vsCMElementClass:
					lastNode = AddXMLDocNode(xdoc, "remarks", lastNode);
					lastNode = AddXMLDocNode(xdoc, "example", lastNode);
					break;
				case vsCMElement.vsCMElementFunction:
					//add returnvalue node
					if ((returnType != null)) {
						returnNode = AddReturnValue(xdoc, returnType, lastNode);
						if (object.ReferenceEquals(returnNode, lastNode)) {
							returnNode = null;
						} else {
							lastNode = returnNode;
						}
					}
					break; // TODO: might not be correct. Was : Exit Select
			}


			//Add documentation for exceptions.
			//Not on classes or interfaces or likewize
			switch (elem.Kind) {
				case vsCMElement.vsCMElementClass:
					break;
				case vsCMElement.vsCMElementInterface:
					break;
				case vsCMElement.vsCMElementEnum:
					break; // TODO: might not be correct. Was : Exit Select

					break;
				default:
					lastNode = AutoDocumentExceptions(elem, xdoc, lastNode);
					break; // TODO: might not be correct. Was : Exit Select

					break;
			}

			//' only update documentation if summaryNode is empty
			if (summaryNode.InnerText.Trim().Length == 0) {
				string summaryText = "";
				string returnText = "";
				string parentFullName = GetParentFullName(elem);
				string seeClassType = GetSeeXmlDocTag(parentFullName);
				string boldClassName = GetHtmlTagType(parentFullName, "b");

				switch (elem.Kind) {
					//Interface documentation
					case vsCMElement.vsCMElementInterface:
						summaryText = "The " + elem.Name + " interface ";
						break;
					//Delegate documentation
					case vsCMElement.vsCMElementDelegate:
						m = Regex.Match(elem.Name, "^(.+)EventHandler$");
						if (m.Success & parameters.Count == 2) {
							summaryText = "Represents a method that will handle a " + m.Groups[1].Value + "Event";
							CodeParameter p1 = (CodeParameter)parameters.Item(1);
							CodeParameter p2 = (CodeParameter)parameters.Item(2);
							paramDocNode = GetParameterXmlNode(xdoc, p1.Name);
							paramDocNode.InnerXml = "The source of the event";
							paramDocNode = GetParameterXmlNode(xdoc, p2.Name);
							paramDocNode.InnerXml = "A " + GetSeeXmlDocTagByType(p2.Type) + " that contains the event data";
						}
						break;
					//Variable documentation
					case vsCMElement.vsCMElementVariable:
						CodeVariable cv = (CodeVariable)elem;
						CodeElement parentElement = (CodeElement)cv.Parent;
						switch (parentElement.Kind) {
							case vsCMElement.vsCMElementEnum:
								summaryText = elem.Name;
								break;
							default:
								if (cv.IsConstant) {
									summaryText = "Get";
								} else {
									summaryText = "Get/Set";
								}
								try {
									//for interface members the cp.Parent doesn't yield the interface, we get an exception instead
									summaryText += "s the " + elem.Name + " of the " + seeClassType;

								} catch {
									//For interfaces, try to retrieve the name of the class from the filename.
									//summaryText += "s the " + elem.Name + " of the " + Strings.Left(DTE.ActiveDocument.Name, Strings.Len(DTE.ActiveDocument.Name) - 3);
									summaryText += "s the " + elem.Name + " of the " + DTE.ActiveDocument.Name.Substring(0, DTE.ActiveDocument.Name.Length - 3);
								}
								break;
						}

						break;
					case vsCMElement.vsCMElementProperty:
						CodeProperty cp = (CodeProperty)elem;

						if (summaryNode.InnerText.Trim().Length == 0) {
							try {
								if ((cp.Getter != null)) {
									summaryText += "Get";
								}
							} catch {
							}

							try {
								if ((cp.Setter != null)) {
									if (summaryText.Length > 0) {
										summaryText += "/";
									}
									summaryText += "Set";
								}
							} catch {
							}

							//2006-10-09 now support explicit generic interface indexers
							if (Regex.IsMatch(cp.Name, "^(?:|[^\\.]+\\.)this$")) {
								summaryText += "s the " + GetSeeXmlDocTagByType(cp.Type) + " item identified by the given arguments of the " + cp.Parent.Name;

							} else {
								try {
									//2006-10-09 now support explicit generic interface indexers
									//for interface members the cp.Parent doesn't yield the interface, we get an exception instead
									summaryText += "s the " + HTMLEncodeText(cp.Name) + " of the " + cp.Parent.Name;
								} catch {
									//2006-10-09 now support explicit generic interface indexers
									//For interfaces, try to retrieve the name of the class from the filename.
									summaryText += "s the " + HTMLEncodeText(cp.Name) + " of the " + DTE.ActiveDocument.Name.Substring(0, DTE.ActiveDocument.Name.Length - 3);
								}

							}
							summaryNode.InnerXml = Environment.NewLine + summaryText + Environment.NewLine;

						}

						break;
					case vsCMElement.vsCMElementFunction:
						CodeFunction2 cf = (CodeFunction2)elem;
						switch (cf.FunctionKind) {
							case vsCMFunction.vsCMFunctionConstructor:
								//Document constructors
								summaryText += "Initializes a new instance of the " + boldClassName + " class.";
								break;
							case vsCMFunction.vsCMFunctionDestructor:
								//Document finalizers
								summaryText += "Releases unmanaged resources and performs other cleanup operations before " + Environment.NewLine;
								summaryText += "the " + boldClassName + " is reclaimed by garbage collection.";
								break;
							default:
								//Format methods
								if (elem.Name.EndsWith("Format") & cf.Parameters.Count == 2) {
									CodeParameter p1 = (CodeParameter)cf.Parameters.Item(1);
									CodeParameter p2 = (CodeParameter)cf.Parameters.Item(2);
									if (p1.Name.Equals("format")) {
										paramDocNode = GetParameterXmlNode(xdoc, p1.Name);
										paramDocNode.InnerXml = "A string containing zero or more format specifications.";
									}
									if (p2.Type.AsString.Equals("object[]")) {
										paramDocNode = GetParameterXmlNode(xdoc, p2.Name);
										paramDocNode.InnerXml = "An array of objects to format.";
									}
								}
								//Check if we have a returnNode
								if (returnNode == null) {
									if (IsMethodsName(elem.Name, "Dispose")) {
										//Document summary of Dispose methods
										summaryText += "Releases the resources used by the " + boldClassName + ".";
										if (cf.Parameters.Count == 1) {
											CodeParameter p1 = (CodeParameter)cf.Parameters.Item(1);
											if (p1.Name == "disposing") {
												paramDocNode = GetParameterXmlNode(xdoc, p1.Name);
												paramDocNode.InnerXml = "Set to <b>true</b> to release both managed and unmanaged resources; <b>false</b> to release only unmanaged resources.";
											}
										}

									} else {
										if (elem.Name.IndexOf("_", 1) > 0 & cf.Parameters.Count == 2) {
											CodeParameter p1 = (CodeParameter)cf.Parameters.Item(1);
											CodeParameter p2 = (CodeParameter)cf.Parameters.Item(2);
											int lastHyphen = cf.Name.LastIndexOf("_");

											string controlName = cf.Name.Substring(0, lastHyphen);
											string eventName = cf.Name.Substring(lastHyphen + 1);

											//we have a event method
											if ((!(p1.Type.CodeType == null | p1.Type.CodeType == null))) {

												if (controlName.Length > 0 & eventName.Length > 0 & p1.Type.CodeType.Name.Equals("Object") & p2.Type.CodeType.Name.EndsWith("Args")) {
													summaryText += "This method is called when the " + controlName + "'s " + eventName + " event has been fired.";
													paramDocNode = GetParameterXmlNode(xdoc, p1.Name);
													paramDocNode.InnerXml = "The " + GetSeeXmlDocTag("object") + " that fired the event.";
													paramDocNode = GetParameterXmlNode(xdoc, p2.Name);
													paramDocNode.InnerXml = "The " + GetSeeXmlDocTagByType(p2.Type) + " of the event.";
												}
											}

										}
									}
								} else {
									//populate returnNode as well
									if (IsMethodsName(cf.Name, "ToString")) {
										//Document summary of ToString methods
										summaryText += "Returns a " + GetSeeXmlDocTag("string") + " that represents the current " + seeClassType + ".";
										//Document return value
										returnNode.InnerXml = "A " + GetSeeXmlDocTag("string") + " that represents the current " + seeClassType + ".";
									} else if (IsMethodsName(cf.Name, "Equals")) {
										//Document Equal methods with one parameter
										if (cf.Parameters.Count == 1) {
											CodeParameter param = (CodeParameter)cf.Parameters.Item(1);
											string seeParamType = GetSeeXmlDocTagByType(param.Type);

											summaryText += "Determines whether the specified " + seeParamType + " is equal to the current " + Environment.NewLine;
											summaryText += boldClassName + ".";
											//Document return value
											returnNode.InnerXml = "true if the specified " + seeParamType + " is equal to the current " + boldClassName + ";" + Environment.NewLine + "otherwise, false.";

											paramDocNode = GetParameterXmlNode(xdoc, param.Name);
											paramDocNode.InnerXml = "The " + seeParamType + " to compare with the current " + seeClassType + ".";

										} else if (cf.Parameters.Count == 2) {
											CodeParameter param = (CodeParameter)cf.Parameters.Item(1);
											CodeParameter param2 = (CodeParameter)cf.Parameters.Item(2);
											string paramTypeName = GetTypeName(param.Type);

											string param2TypeName = GetTypeName(param2.Type);
											if (paramTypeName.Equals(param2TypeName)) {
												string seeParamType = GetSeeXmlDocTag(paramTypeName);
												string paramName = param.Name;
												string param2Name = param2.Name;


												summaryText += "Determines whether the specified " + seeParamType + " instances are considered equal.";
												//Document return value
												returnNode.InnerXml = "true if " + GetHtmlTagType(paramName, "i") + " is the same instance as " + GetHtmlTagType(param2Name, "i") + " " + Environment.NewLine;
												returnNode.InnerXml += "or  if both are null references or if <c>" + paramName + ".Equals(" + param2Name + ")</c> returns true; otherwise, false.";

												paramDocNode = GetParameterXmlNode(xdoc, param.Name);
												paramDocNode.InnerXml = "The first " + seeParamType + " to compare.";
												paramDocNode = GetParameterXmlNode(xdoc, param2.Name);
												paramDocNode.InnerXml = "The second " + seeParamType + " to compare.";
											}
										}
									} else if (IsMethodsName(cf.Name, "GetHashCode")) {
										//Document GetHashCode methods
										summaryText += "Serves as a hash function for a particular type, suitable for use in hashing algorithms and data structures like a hash table.";
										//Document return value
										returnNode.InnerXml = "A hash code for the current " + seeClassType + ".";
									} else if (IsMethodsName(cf.Name, "Clone")) {
										summaryText += "Creates a new object that is a copy of the current instance.";
										returnNode.InnerXml = "A new object that is a copy of this instance.";
									} else if (IsMethodsName(cf.Name, "Contains")) {
										//Document Equal methods with one parameter
										if (cf.Parameters.Count > 0) {
											CodeParameter param = (CodeParameter)cf.Parameters.Item(1);
											string paramTypeName = GetTypeName(param.Type);
											string seeParamType = GetSeeXmlDocTag(paramTypeName);


											summaryText += "Returns a value indicating whether the specified " + seeParamType + Environment.NewLine;
											summaryText += " is contained in the " + seeClassType + ".";
											returnNode.InnerXml = "<b>true</b> if the " + GetHtmlTagType(paramTypeName, "i") + " parameter is a member " + Environment.NewLine;
											returnNode.InnerXml += "of the " + seeClassType + "; otherwise, <b>false</b>.";

											paramDocNode = GetParameterXmlNode(xdoc, param.Name);
											paramDocNode.InnerXml = "The " + seeParamType + " to locate in the " + Environment.NewLine + seeClassType + ".";
										}

										//IsWordWordWord
									} else if (cf.Name.Length > 2 & cf.Name.StartsWith("Is") & IsTypeNamed(cf.Type, "Boolean")) {
										string[] words = SplitOnPascalCasing(elem.Name);
										returnNode.InnerXml = "True if ";
										if (words.Length > 2) {
											returnNode.InnerXml += string.Join(" ", words, 1, words.Length - 2).ToLower();
										} else {
											returnNode.InnerXml += "it";
										}
										returnNode.InnerXml += " is " + words[words.Length - 1].ToLower() + ", otherwise false.";
									} else if (IsMethodsName(cf.Name, "CompareTo") & cf.Parameters.Count == 1 & IsTypeNamed(cf.Type, "Int32")) {
										summaryText += "Compares the current instance with another object of the same type.";
										CodeParameter param = (CodeParameter)cf.Parameters.Item(1);
										paramDocNode = GetParameterXmlNode(xdoc, param.Name);
										string paramTypeName = GetTypeName(param.Type);
										string seeParamType = GetSeeXmlDocTag(paramTypeName);
										string iParamName = GetHtmlTagType(param.Name, "i");

										paramDocNode.InnerXml = "The " + seeParamType + " to compare with this instance." + Environment.NewLine;
										returnText += "A 32-bit signed integer that indicates the relative order " + Environment.NewLine;
										returnText += "of the objects being compared. The return value has these meanings: " + Environment.NewLine;
										returnText += "<table>" + Environment.NewLine;
										returnText += "<tr><th>Value</th><th>Meaning</th></tr>" + Environment.NewLine;
										returnText += "<tr><td>Less than zero</td><td>This instance is less than " + iParamName + ".</td></tr>" + Environment.NewLine;
										returnText += "<tr><td>Zero</td><td>This instance is equal to " + iParamName + ".</td></tr>" + Environment.NewLine;
										returnText += "<tr><td>Greater than zero</td><td>This instance is greater than " + iParamName + ".</td></tr>" + Environment.NewLine;
										returnText += "</table>" + Environment.NewLine;
									} else if (IsMethodsName(cf.Name, "Compare") & cf.Parameters.Count == 2 & IsTypeNamed(cf.Type, "Int32")) {
										summaryText += "Compares two objects and returns a value indicating whether one is less than," + Environment.NewLine + "equal to, or greater than the other.";
										CodeParameter p1 = (CodeParameter)cf.Parameters.Item(1);
										CodeParameter p2 = (CodeParameter)cf.Parameters.Item(2);
										paramDocNode = GetParameterXmlNode(xdoc, p1.Name);
										paramDocNode.InnerXml = "The first object to compare.";
										paramDocNode = GetParameterXmlNode(xdoc, p2.Name);
										paramDocNode.InnerXml = "The second object to compare.";

										returnText += "A 32-bit signed integer that indicates the relative order " + Environment.NewLine;
										returnText += "of the objects being compared. The return value has these meanings: " + Environment.NewLine;
										returnText += "<table>" + Environment.NewLine;
										returnText += "<tr><th>Value</th><th>Meaning</th></tr>" + Environment.NewLine;
										returnText += "<tr><td>Less than zero</td><td>x less than y.</td></tr>" + Environment.NewLine;
										returnText += "<tr><td>Zero</td><td>x is equal to y.</td></tr>" + Environment.NewLine;
										returnText += "<tr><td>Greater than zero</td><td>x is greater than y.</td></tr>" + Environment.NewLine;
										returnText += "</table>" + Environment.NewLine;


									}
								}
								break;
						}

						break;
				}

				if (string.IsNullOrEmpty(summaryText)) {
					summaryText = " ";
				}
				if (returnText.Length > 0) {
					returnNode.InnerXml = returnText;
				}
				summaryNode.InnerXml = Environment.NewLine + summaryText + Environment.NewLine;
			}
			return summaryNode;
		}
		#endregion
		private XmlNode AutoDocumentExceptions(CodeElement elem, XmlDocument xdoc, XmlNode lastNode) {
			//Add documentation for exceptions.
			string text = GetElementContent(elem);
			MatchCollection ms = Regex.Matches(text, "throw\\s+new\\s+(.*?Exception)\\s*\\((.*?)\\)\\s*;", RegexOptions.Singleline);
			//Dim argumentNullExceptionNode As XmlNode
			//Dim argumentNullExceptionMessage As String = String.Empty
			//Dim argumentOutOfRangeExceptionNode As XmlNode
			//Dim argumentOutOfRangeExceptionMessage As String = String.Empty
			Dictionary<XmlNode, string> exceptionDict = new Dictionary<XmlNode, string>();

			foreach (Match m1 in ms) {

				if (m1.Success) {
					//make sure that the exception node is in place
					string ex = m1.Groups[1].Value;
					string @params = m1.Groups[2].Value;

					XmlNode node = EnsureNode(xdoc, "/doc/exception[@cref='" + ex + "']");
					if ((!object.ReferenceEquals(lastNode, node)) & (!object.ReferenceEquals(lastNode.NextSibling, node))) {
						SwapNodes(lastNode.NextSibling, node);
					}
					lastNode = node;

					//Try to autodocument the exceptions
					Match m = default(Match);
					if (string.IsNullOrEmpty(node.InnerXml) | exceptionDict.ContainsKey(node)) {
						string text2 = null;

						if (exceptionDict.ContainsKey(node)) {
							text2 = exceptionDict[node];
						} else {
							text2 = string.Empty;
							exceptionDict.Add(node, text2);
						}
						m = Regex.Match(@params, "^\\s*\"([^\"]+)\"(?:\\s*,\\s*\"([^\"]+)\")?");

						if (m.Success) {
							string tt = string.Empty;
							switch (ex) {
								case "ArgumentNullException":
									tt = "<paramref name=\"" + m.Groups[1].Value + "\"/> is null";
									break;
								case "ArgumentOutOfRangeException":
									tt = "<paramref name=\"" + m.Groups[1].Value + "\"/> is out of range";
									break;
								default:
									if (m.Groups[2].Value.Length == 0) {
										tt = m.Groups[1].Value.Substring(0, 1).ToLower() + m.Groups[1].Value.Substring(1);
									}
									break;
							}


							if (!string.IsNullOrEmpty(tt)) {
								//add message to exception details
								if (m.Groups.Count > 1 & m.Groups[2].Value.Length > 0) {
									tt += " (<em>" + m.Groups[2].Value + "</em>)";
								}

								if (text2.Length > 0) {
									text2 += Environment.NewLine + " or if ";
								} else {
									text2 = "If ";
								}
								text2 += tt;
							}
						}

						exceptionDict[node] = text2;

					}
				}
			}
			//Set the autodocumentation message for each exception.
			foreach (KeyValuePair<XmlNode, string> pair in exceptionDict) {
				if (!string.IsNullOrEmpty(pair.Value)) {
					pair.Key.InnerXml = pair.Value;
					if (!pair.Value.EndsWith(".")) {
						pair.Key.InnerXml += ".";
					}
				}
			}
			return lastNode;
		}
		#region "Private Function AddParameters(ByVal xdoc As XmlDocument, ByVal parameters As CodeElements, ByVal lastNode As XmlNode) As XmlNode"
		/// <summary>
		///	 Adds the parameters to the xml document.
		/// </summary>
		/// <param name="xdoc" >The xml document</param>
		/// <param name="parameters" >The collection of parameters</param>
		/// <param name="lastNode" >The node to place the created/found node after</param>
		/// <returns ></returns>
		private XmlNode AddParameters(XmlDocument xdoc, CodeElements parameters, XmlNode lastNode) {
			//add parameter nodes
			foreach (CodeParameter cp in parameters) {
				XmlNode node = GetParameterXmlNode(xdoc, cp.Name);
				if ((!object.ReferenceEquals(lastNode.NextSibling, node))) {
					SwapNodes(lastNode.NextSibling, node);
				}
				lastNode = node;
			}
			//'todo: remove old params
			return lastNode;
		}
		#endregion
		#region "Private Function GetParameterXmlNode(ByVal xdoc As XmlDocument, ByVal name As String) As XmlNode"
		/// <summary>
		/// Gets the node containing the supplied parameter documentation
		/// </summary>
		/// <param name="xdoc">The Xmldocument containing the nodes</param>
		/// <param name="name">The name of the parameter</param>
		/// <returns></returns>
		private XmlNode GetParameterXmlNode(XmlDocument xdoc, string name) {
			return EnsureNode(xdoc, "/doc/param[@name='" + name + "']");
		}
		#endregion
		#region "Private Function RemoveEmptyElements(ByVal xdoc As XmlDocument)"
		/// <summary>
		/// Removes all empty child elements from the /doc element
		/// </summary>
		/// <param name="xdoc">The xml document</param>
		private object RemoveEmptyElements(XmlDocument xdoc) {
			object functionReturnValue = null;
			XmlNode rootNode = xdoc.SelectSingleNode("/doc");
			if (rootNode == null) {
				return functionReturnValue;
			}
			XmlNodeList nodes = xdoc.SelectNodes("/doc/*");

			foreach (XmlNode node in nodes) {
				//'skip inheritdoc nodes
				if (node.Name == "inheritdoc") {
					continue;
				}

				if (node.InnerXml.Trim().Length == 0) {
					rootNode.RemoveChild(node);
				}
			}
			return functionReturnValue;

		}
		#endregion
		#region "Private Function AddReturnValue(ByVal xdoc As XmlDocument, ByVal type As CodeTypeRef, ByVal lastNode As XmlNode) As XmlNode"
		/// <summary>
		///	 Adds the returnvalue parameter to the xmldocument
		/// </summary>
		/// <param name="xdoc" >The xml document</param>
		/// <param name="type" >The type of returnvalue</param>
		/// <param name="lastNode" >The node to place the created/found node after</param>
		/// <returns ></returns>
		private XmlNode AddReturnValue(XmlDocument xdoc, CodeTypeRef type, XmlNode lastNode) {
			//add returnvalue node if return type is not null
			if (type.TypeKind != vsCMTypeRef.vsCMTypeRefVoid) {
				return AddXMLDocNode(xdoc, "returns", lastNode);
			}
			return lastNode;
		}
		#endregion
		#region "Private Function AddXMLDocNode(ByVal xdoc As XmlDocument, ByVal name As String, ByVal lastNode As XmlNode) As XmlNode"
		/// <summary>
		/// Adds the supplied xml doc node after the supplied lastNode. 
		/// if the node already exists it is placed after the supplied lastNode.
		/// </summary>
		/// <param name="xdoc"></param>
		/// <param name="name"></param>
		/// <param name="lastNode"></param>
		/// <returns></returns>
		private XmlNode AddXMLDocNode(XmlDocument xdoc, string name, XmlNode lastNode) {
			XmlNode node = EnsureNode(xdoc, "/doc/" + name);
			if ((!object.ReferenceEquals(lastNode.NextSibling, node))) {
				SwapNodes(lastNode.NextSibling, node);

			}
			lastNode = node;
			return lastNode;

		}
		#endregion
		#region "Private Function PrepareXMLDocumentation(ByVal element As CodeElement) As XmlDocument"
		/// <summary>
		/// Prepares the XML documentation for the supplied element
		/// </summary>
		/// <param name="element">The <see cref="CodeElement"/> prepare</param>
		/// <returns>The prepared <see cref="XmlDocument"/> </returns>
		protected XmlDocument PrepareXMLDocumentation(CodeElement element) {
			System.Reflection.PropertyInfo pInfo = GetPropertyInfo(element, "DocComment");
			if (pInfo == null) {
				return null;
			}
			string xml = Convert.ToString(pInfo.GetValue(element, null));

			XmlDocument xdoc = new XmlDocument();

			if (xml == null | xml.Length == 0) {
				xml = "<doc></doc>";
			} else {
				if (IsBasic) {
					//strip all whitespaces before a start tag...
					//ignore whitespaces inside tags...
					xml = Regex.Replace(xml, "\\s*<(?!/)", "<");
					xml = "<doc>" + xml.Trim() + "</doc>";
				}
			}
			xdoc.PreserveWhitespace = true;

			xdoc.LoadXml(xml);

			//EnsureNode(xdoc, "remarks", vbCrLf & " " & vbCrLf, True)

			return xdoc;

		}
		#endregion
		#region "Private Function EnsureNode(ByVal xdoc As XmlDocument, ByVal path As String, Optional ByVal defaultValue As String = "", Optional ByVal firstNode As Boolean = False) As XmlNode"
		/// <summary>
		/// Makes sure that the supplied node exists in the document
		/// </summary>
		/// <param name="xdoc" ></param>
		/// <param name="path" ></param>
		/// <param name="defaultValue" ></param>
		/// <param name="firstNode" ></param>
		/// <returns ></returns>
		protected XmlNode EnsureNode(XmlDocument xdoc, string path, string defaultValue = "", bool firstNode = false) {
			XmlNode node = CreateElement(xdoc, path);


			//do we need to place it at the front?
			if (firstNode & (!object.ReferenceEquals(xdoc.FirstChild.FirstChild, node))) {
				SwapNodes(xdoc.FirstChild.FirstChild, node);
			}
			if ((defaultValue != null)) {
				if (node.InnerText.Length == 0 & defaultValue.Length > 0) {
					node.InnerXml = defaultValue;
				}
			}
			return node;

		}
		#endregion
		private bool HasNode(XmlDocument xdoc, string path) {
			XmlNode node = CreateElement(xdoc, path, false);
			if (node == null) {
				return false;
			}
			return true;
		}
		#region "Private Function SwapNodes(ByVal xdoc As XmlDocument, ByVal n1 As XmlNode, ByVal n2 As XmlNode)"
		/// <summary>
		///	 swaps two nodes places. Will only work if the nodes have the same parent.
		/// </summary>
		/// <param name="n1" ></param>
		/// <param name="n2" ></param>
		/// <returns ></returns>
		private object SwapNodes(XmlNode n1, XmlNode n2) {
			object functionReturnValue = null;
			if ((!object.ReferenceEquals(n1.ParentNode, n2.ParentNode))) {
				return functionReturnValue;
			}
			XmlNode tmp = n1.Clone();
			XmlNode parent = n1.ParentNode;

			parent.ReplaceChild(tmp, n1);
			parent.ReplaceChild(n1, n2);
			parent.ReplaceChild(n2, tmp);
			return functionReturnValue;

		}
		#endregion
		private CodeElement GetParent(CodeElement elem) {
			var cm = elem as CodeFunction;
			if (cm != null)
				return (cm.Parent as CodeElement);
			var cv = elem as CodeVariable;
			if (cv != null)
				return cv.Parent as CodeElement;
			var cc = elem as CodeClass;
			if (cc != null)
				return cc.Parent as CodeElement;
			var ca = elem as CodeAttribute;
			if (ca != null)
				return ca.Parent as CodeElement;
			var cd = elem as CodeDelegate;
			if (cd != null)
				return cd.Parent as CodeElement;
			var ce = elem as CodeEnum;
			if (ce != null)
				return ce.Parent as CodeElement;
			var cev = elem as CodeEvent;
			if (cev != null)
				return cev.Parent as CodeElement;
			var ci = elem as CodeImport;
			if (ci != null)
				return ci.Parent as CodeElement;
			var cmo = elem as CodeModel;
			if (cmo != null)
				return cmo.Parent as CodeElement;
			var cn = elem as CodeNamespace;
			if (cn != null)
				return cn.Parent as CodeElement;
			var cp = elem as CodeParameter;
			if (cp != null)
				return cp.Parent as CodeElement;
			var cpr = elem as CodeProperty;
			if (cpr != null)
				return cpr.Parent as CodeElement;
			var cs = elem as CodeStruct;
			if (cs != null)
				return cs.Parent as CodeElement;
			var ct = elem as CodeType;
			if (ct != null)
				return ct.Parent as CodeElement;
			var ctr = elem as CodeTypeRef;
			if (ctr != null)
				return ctr.Parent as CodeElement;
			return null;
		}
		private string GetParentFullName(CodeElement elem) {
			try {
				var parent = GetParent(elem);
				if (parent != null)
					return parent.FullName;
			} catch {
			}
			//Some items don't have the Parent property, (like interface properties). Then we can use the full name of the elem and only remove the last part (ie. the poperty name)
			return elem.FullName.Substring(1, elem.FullName.LastIndexOf(".", System.StringComparison.Ordinal) - 1);

		}
		#region "Private Function GetDocumentationElement() As CodeElement"
		/// <summary>
		/// Gets the element that contains the documentation.
		/// </summary>
		/// <returns>The found element</returns>
		private CodeElement GetDocumentationElement() {
			CodeElement element1 = null;
			if ((DTE.ActiveWindow.Type != vsWindowType.vsWindowTypeDocument)) {
				return element1;
			}
			if ((!IsCSharp & !IsBasic)) {
				return element1;
			}
			TextSelection selection1 = (TextSelection)DTE.ActiveWindow.Document.Selection;
			Point startPoint = GetPoint(selection1.ActivePoint);

			try {
				element1 = GetElement();
				if (element1.Kind == vsCMElement.vsCMElementVariable) {
					return element1; // TODO: might not be correct. Was : Exit Try
				}


				if (IsBasic) {
					//Since basic Get set properties are flagged as ElementFunctions, we need to
					//find the overlying property
					CodeElement element2 = element1;
					while (element2.Kind == vsCMElement.vsCMElementFunction & (element2.Name == "Get" | element2.Name == "Set")) {
						selection1.MoveToPoint(element2.StartPoint);
						selection1.LineUp();
						element2 = GetElement();
					}
					if (element2.Kind == vsCMElement.vsCMElementProperty) {
						element1 = element2;
						return element1; // TODO: might not be correct. Was : Exit Try
					}
				}


				MoveToPoint(startPoint);

			} catch (Exception ex) {
				MoveToPoint(startPoint);
			}
			return element1;
		}
		#endregion
		
		#region "Private Function GetHtmlTagType(ByVal type As String, ByVal tag As String) As String"
		/// <summary>
		/// Gets a html tag with the supplied type as content
		/// </summary>
		/// <param name="type">Contents of the tag</param>
		/// <param name="tag">Name of the tag</param>
		/// <returns>The constructed tag</returns>
		private string GetHtmlTagType(string type, string tag) {
			//remove namespace from type.
			type = Regex.Replace(type, "^.*\\.", "", RegexOptions.None);

			type = type.Replace("<", "&lt;").Replace(">", "&gt;");
			return "<" + tag + ">" + type + "</" + tag + ">";
		}
		#endregion
		
		#region "Private Function HTMLEncodeText(ByVal text As String) As String"
		/// <summary>
		/// Encodes the supplied text to html
		/// </summary>
		/// <param name="text">The <see cref="String"/>  to encode</param>
		/// <returns>The encoded <see cref="String"/> </returns>
		private string HTMLEncodeText(string text) {
			text = Regex.Replace(text, "&", "&amp;", RegexOptions.None);
			text = Regex.Replace(text, "<", "&lt;", RegexOptions.None);
			text = Regex.Replace(text, ">", "&gt;", RegexOptions.None);
			text = Regex.Replace(text, "\"", "&quot;", RegexOptions.None);
			return text;
		}
		#endregion
		#region "Private Function IsMethodsName(ByVal text As String, ByVal name As String) As Boolean"
		/// <summary>
		/// Checks if the supplied text contains a method name of the supplied name. (inlcuding explicit interface member names)
		/// </summary>
		/// <param name="text">The text to check</param>
		/// <param name="name">The name to look for</param>
		/// <returns>True if the name was found, false otherwize</returns>
		private bool IsMethodsName(string text, string name) {
			return Regex.IsMatch(text, "(?:\\w+\\.)?" + name);
		}
		#endregion
		#region "Private Function SplitOnPascalCasing(ByVal text As String) As String()"
		/// <summary>
		/// Splits a string on pascal casing
		/// </summary>
		/// <param name="text">The text to split</param>
		/// <returns>an array of strings</returns>
		private string[] SplitOnPascalCasing(string text) {
			string[] words = Regex.Split(text, "([A-Z][^A-Z]*)");
			var l = new List<string>();
			foreach (string w in words) {
				if (w.Length > 0) {
					l.Add(w);

				}
			}
			return l.ToArray();

		}
		#endregion
		#region "Private Function IsTypeNamed(ByVal type As CodeTypeRef, ByVal name As String) As Boolean"
		/// <summary>
		/// Checks if the supplied type has the supplied name
		/// </summary>
		/// <param name="type">The <see cref="CodeTypeRef"/> to check</param>
		/// <param name="name">The name of the type to compare with</param>
		/// <returns>True if type is named as the supplied name, otherwise false.</returns>
		private bool IsTypeNamed(CodeTypeRef type, string name) {
			if (name.EndsWith("[]")) {
				if (type.TypeKind == vsCMTypeRef.vsCMTypeRefArray) {
					return IsTypeNamed(type.ElementType, name.Substring(0, name.Length - 2));
				}
			}
			if (type.TypeKind == vsCMTypeRef.vsCMTypeRefArray) {
				return false;
			}
			//In case of Generic types..
			if (type.TypeKind == vsCMTypeRef.vsCMTypeRefOther) {
				return false;
			}
			if ((type.CodeType != null)) {
				return name.Equals(type.CodeType.Name);
			}
			return name.Equals(type.AsString);
		}
		#endregion

		private string GetElementContent(CodeElement elem) {
			TextSelection sel = (TextSelection)DTE.ActiveWindow.Document.Selection;
			sel.MoveToPoint(elem.StartPoint);
			sel.MoveToPoint(elem.EndPoint, true);
			return sel.Text;
		}
		#region "Private Function CreateElement(ByVal doc As XmlDocument, ByVal elementPath As String) As XmlNode"
		/// <summary>
		/// Creates an element at the supplied path. If the parent elements doesn't exists, they are created aswell.
		/// If the element already exists, it will be returned.
		/// </summary>
		/// <param name="doc">The document that the element should be created in</param>
		/// <param name="elementPath">The absolute path to the element</param>
		/// <returns>The created (or found) node</returns>
		private XmlNode CreateElement(XmlDocument doc, string elementPath, bool create = true) {
			if (elementPath == null) {
				throw new ArgumentNullException("elementPath");
			}
			string[] paths = elementPath.Split('/');
			XmlNode lastNode = doc;
			int i = 0;

			for (i = 0; i <= paths.Length - 1; i++) {
				if (paths[i].Length > 0) {
					XmlNode node = lastNode.SelectSingleNode(paths[i]);
					if (node == null) {
						//'If create is of, return nothing
						if (!create) {
							return null;
						}


						if (paths[i].StartsWith("@")) {
							XmlAttribute att = doc.CreateAttribute(paths[i].Substring(1));
							lastNode.Attributes.Append(att);
							node = att;
						} else {
							Match m = Regex.Match(paths[i], "([^\\[]*)");


							if (m.Success) {
								string elementName = m.Groups[0].Value;
								node = doc.CreateElement(elementName);
								lastNode.AppendChild(node);
								m = Regex.Match(paths[i], "\\[(.*)\\]");
								if (m.Success) {
									//parse arguments and add nesseccary attributes
									string arg = m.Groups[0].Value;
									//not? @name='?value'? cond?
									string pattern = "((?:not)?)\\s?@([^\\\\s=]+)\\s?=\\s?((?:\\'[^\\']*\\')|(?:[^\\s]*))\\s?([^\\s\\]]+)?";
									m = Regex.Match(arg, pattern, RegexOptions.IgnoreCase);

									while (m.Success) {
										string strNot = m.Groups[1].Value;
										string attName = m.Groups[2].Value;
										string attVal = m.Groups[3].Value;
										if (attVal[0] == '\'' & attVal.Length >= 2) {
											attVal = attVal.Substring(1, attVal.Length - 2);
										}

										string cond = m.Groups[4].Value;
										switch (cond.ToLower().Trim()) {
											case "and":
												break; // TODO: might not be correct. Was : Exit Select

												break;
											case "or":
												break; // TODO: might not be correct. Was : Exit Select

												break;
											case "":
												break; // TODO: might not be correct. Was : Exit Select

												break;
											default:
												lastNode.RemoveChild(node);
												throw new ArgumentException("The elementpath doesn't support condition " + cond + " in \\" + arg + " / ", "elementPath");
										}
										//set the attribute (if it should be there)
										if (strNot.Length == 0) {
											SetAttribute(node, attName, attVal);

										}

										m = m.NextMatch();
									}
								}
							}
						}
					}
					lastNode = node;
				}
			}
			return lastNode;
		}
		#endregion
		#region "Private Function SetAttribute(ByVal node As XmlNode, ByVal attrName As String, ByVal attrValue As String)"
		/// <summary>
		/// Sets an attribute value in the the supplied node
		/// </summary>
		/// <param name="node">The node to create the attribute in</param>
		/// <param name="attrName">The name of the attribute</param>
		/// <param name="attrValue">The value of the attribute</param>
		/// <returns></returns>
		private void SetAttribute(XmlNode node, string attrName, string attrValue) {
			XmlAttribute newAttributeNode = node.OwnerDocument.CreateAttribute(attrName);
			newAttributeNode.Value = attrValue;
			node.Attributes.SetNamedItem(newAttributeNode);
		}
		#endregion
		private void DocumentAndRegionizeHtml(bool regionize) {
			Point p = GetPoint();

			if (Selection.FindPattern(@"^\s*function", (int)vsFindOptions.vsFindOptionsBackwards)) {
				Selection.StartOfLine();
				Selection.Collapse();

				Point p2 = GetPoint();
				//TODO:unfinished business
			}

		}

		/// <summary>
		/// Enters a remark on the format yyyy-MM-dd, domain/user:
		/// </summary>
		[Command("Text Editor::Ctrl+Shift+Alt+c")]
		public void EnterCodeRemark() {
			CodeElement element1 = DocumentAndRegionize(false);
			if (element1 == null) {
				return;
			}
			TextSelection selection1 = (TextSelection)DTE.ActiveWindow.Document.Selection;
			Point startPoint = GetPoint(selection1.ActivePoint);

			Type type = GetElementType(element1);
			if (type == null) {
				return;
			}


			try {
				XmlDocument xdoc = PrepareXMLDocumentation(element1);
				if (xdoc == null) {
					MoveToPoint(startPoint);
					return;
				}
				EnsureNode(xdoc, "/doc/summary", " ", true);
				XmlNode remarksNode = EnsureNode(xdoc, "/doc/remarks");
				//Dim remarksNode As XmlNode = CreateElement(xdoc, "/doc/remarks")
				remarksNode.InnerXml += Environment.NewLine + "<para>" + DateTime.Now.ToString("yyyy-MM-dd") + ", " + Environment.UserDomainName + "\\" + Environment.UserName + ": |</para>";
				SaveXMLDocumentation(xdoc, element1);
				SetCursorToNodeText(remarksNode, true, true);

			} catch (Exception exception1) {
				//if anything goes wrong, return to old position
				MoveToPoint(startPoint);
				element1 = null;
				MessageBox.Show(exception1.Message, "Documentator", MessageBoxButtons.OK, MessageBoxIcon.Hand);

			}

		}
	}
}