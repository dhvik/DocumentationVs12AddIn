﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Text.RegularExpressions;
using EnvDTE;

namespace DocumentationVs12AddIn.Commands {
	/// <summary>
	/// Summary description for MiscCommands.
	/// </summary>
	/// <remarks>
	/// 2013-04-08 dan: Created
	/// </remarks>
	public class MiscCommands:CommandBase {
		private Dictionary<string,string> ExtractAttributesList;

		/* *******************************************************************
		 *  Commands
		 * *******************************************************************/
		#region public void KillLine()
		/// <summary>
		/// 
		/// </summary>
		[Command("Text Editor::Ctrl+Shift+k")]
		public void KillLine() {
			if (Selection == null)
				return;
			Selection.StartOfLine();
			Selection.EndOfLine(true);
			Selection.Delete();
		}
		#endregion
		#region public void CreateDocumentationHeading()
		/// <summary>
		///  Creates a documentation Heading for the current line
		/// </summary>
		[Command("Text Editor::Ctrl+Shift+Alt+h")]
		public void CreateDocumentationHeading() {
			var sel = Selection;
			var p = GetPoint();
			sel.SelectLine();
			var text = sel.Text;
			Match m = Regex.Match(text, "^(\\s*)(.*?)\\s*$");
			if (m.Success) {
				string indent = m.Groups[1].Value;
				text = m.Groups[2].Value;

				string data = "";
				if (IsBasic) {
					text = Regex.Replace(text, "^'+", "");
					data += indent + "'' *******************************************************************" + Environment.NewLine;
					data += indent + "''  " + text + Environment.NewLine;
					data += indent + "'' *******************************************************************" + Environment.NewLine;
				} else {
					text = Regex.Replace(text, "^/{2,}", "");
					data += indent + "/* *******************************************************************" + Environment.NewLine;
					data += indent + " *  " + text + Environment.NewLine;
					data += indent + " * *******************************************************************/" + Environment.NewLine;
				}
				sel.Delete();
				sel.Insert(data);
			} else {
				MoveToPoint(p);
			}
		}
		#endregion
		#region public void PasteSeeParameterXmlDocTag()
		/// <summary>
		/// Pastes the See parameter xml document tag in the current documentation section
		/// </summary>
		[Command("Text Editor::Ctrl+Shift+x, Ctrl+Shift+d")]
		public void PasteSeeParameterXmlDocTag() {
			PasteSeeXmlDocParamTag();
		}
		#endregion
		#region public void MakeCData()
		/// <summary>
		/// Makes the selected text CData and adds pre tags around it.
		/// </summary>
		public void MakeCData() {
			if ((DTE.ActiveWindow.Type != vsWindowType.vsWindowTypeDocument)) {
				return;
			}

			TextSelection sel = Selection;
			string text = sel.Text;
			//text = Regex.Replace(text, "&", "&amp;", RegexOptions.None)
			//text = Regex.Replace(text, "<", "&lt;", RegexOptions.None)
			//text = Regex.Replace(text, ">", "&gt;", RegexOptions.None)
			//text = Regex.Replace(text, """", "&quote;", RegexOptions.None)
			sel.Delete();
			sel.Insert("<pre><![CDATA[" + text + "]]></pre>");
		}
		#endregion
		#region public void FixHtmlValidationErrors()
		/// <summary>
		/// Fixes some HTML validation errors.
		/// </summary>
		public void FixHtmlValidationErrors() {
			if ((DTE.ActiveWindow.Type != vsWindowType.vsWindowTypeDocument)) {
				return;
			}

			TextSelection sel = (TextSelection)DTE.ActiveWindow.Document.Selection;
			Point p = GetPoint();

			sel.SelectAll();

			string text = sel.Text;

			//make sure that the html tag contains the xmlns attribute.
			text = EnsureAttribute(text, "html", "xmlns", "http://www.w3.org/1999/xhtml");

			//Make all elements to lowercase
			//make all attribute names to lowercase
			//fix all attributes to be key="Value" (from key = "value", key = 'value', key = value, key
			text = Regex.Replace(text, "<\\/?(\\w+\\b)(?:\\s+([^<>\\s=]+(?:(?:\\s*=\\s*\"[^\"]*\")|(?:\\s*=\\s*'[^']*')|(?:\\s*=\\s*\\S+)|(?:))))*?\\s*\\/?>", FixAttributesMatchEvaluator, RegexOptions.None);
			//text = Regex.Replace(text, "<\/?(\w+\b)(?:\s+(\w+)\s*=\s*""[^""]*"")*\s*\/?>", ToLowerCaseMatchEvaluatorRef, RegexOptions.None)
			//merge empty elements from <elem></elem> to <elem/>
			text = Regex.Replace(text, "<(\\w+)((?:\\s+\\w+\\s*=\\s*\"[^\"]*\")*)\\s*>\\s*<\\/\\1>", "<$1$2/>", RegexOptions.IgnoreCase);
			//make sure that the following elements has a end element.
			text = Regex.Replace(text, "(<(?:link|meta|img|hr|br)(?:\\s+\\w+\\s*=\\s*\"[^\"]*\")*\\s*)>", "$1/>", RegexOptions.IgnoreCase);
			//replace nowrap="whatever" with 
			//TODO
			//add end </li> where missing
			text = AssertEndElement(text, "li", new string[] {
		"ul",
		"ol"
	});
			text = AssertEndElement(text, "p", new string[] {
		"body",
		"h1",
		"h2",
		"h3",
		"h4",
		"h5",
		"h6",
		"dl",
		"hr"
	});
			text = AssertEndElement(text, "dt", new string[] {
		"dl",
		"dd"
	});
			//deny script to have self closing tags convert <script /> to <script></script>
			text = Regex.Replace(text, "(<script\\b[^>]*)/>", "$1></script>", RegexOptions.IgnoreCase | RegexOptions.Singleline);

			text = Regex.Replace(text, "(<\\w+\\b[^>]*style\\s*=\\s*\"[^\"]*)(\"\\s[^>]*)style\\s*=\\s*\"([^\"]*)\"", "$1;$3$2", RegexOptions.IgnoreCase | RegexOptions.Singleline);

			//remove language attributes from div tags
			text = Regex.Replace(text, "(<div\\b[^>]*)\\s+language\\s*=\\s*\"[^\"]*\"([^>]*>)", "$1$2", RegexOptions.IgnoreCase);
			//remove ms_positioning attributes from div tags
			text = Regex.Replace(text, "(<div\\b[^>]*)\\s+ms_positioning\\s*=\\s*\"[^\"]*\"([^>]*>)", "$1$2", RegexOptions.IgnoreCase);
			//remove topmargin attributes from body tags
			text = Regex.Replace(text, "(<body\\b[^>]*)\\s+topmargin\\s*=\\s*\"[^\"]*\"([^>]*>)", "$1$2", RegexOptions.IgnoreCase | RegexOptions.Singleline);
			//add alt attribute on img tags if missing (set it to the same as the src attribute)
			text = EnsureAttribute(text, "img", "alt", "", "src");


			sel.Delete();
			sel.Insert(text);

			//  DTE.ExecuteCommand("Edit.FormatSelection")

			MoveToPoint(p);
			return;
			//Dim tagsToLowerCase As String() = {"HEAD", "LINK", "TITLE", "BODY", "HTML", "DIV", "TABLE", "TR", "TD", "TH", "H1", "H2", "H3", "H4", "H5", "PRE", "P", "A", "EM", "STRONG", "I", "B", "MSHelp\:TOCTitle", "MSHelp\:RLTitle", "MSHelp\:Keyword", "MSHelp\:Attr", "UL", "OL", "LI", "IMG"}
			//For Each tag As String In tagsToLowerCase
			//    MakeHTMLTagLowercase(tag)
			//Next
			//Replace("XMLNS\:MSHelp", "xmlns:mshelp")







			//MakeHTMLTagLowercase("DIV")
		}
		#endregion
		#region public void EnterCodeComment()
		/// <summary>
		/// Enters a comment on the format yyyy-MM-dd, domain/user: 
		/// </summary>
		[Command("Text Editor::Ctrl+Shift+c")]
		public void EnterCodeComment() {
			TextSelection objSelection = Selection;
			objSelection.Collapse();
			objSelection.StartOfLine();
			objSelection.EndOfLine(true);
			string indent = "";
			Match m = Regex.Match(objSelection.Text, "^(\\s*)");
			if ((m.Success)) {
				indent = m.Groups[1].Value;
			}
			objSelection.Collapse();
			objSelection.StartOfLine();

			string comment = null;
			string commentLineStart = "// ";
			if (IsBasic) {
				commentLineStart = "' ";
			}

			comment = indent + commentLineStart + DateTime.Now.ToString("yyyy-MM-dd") + ", " + Environment.UserDomainName + "\\" + Environment.UserName + ": ";
			objSelection.Insert(comment + Environment.NewLine);
			objSelection.LineUp();
			objSelection.EndOfLine();


		}
		#endregion
		#region public void FormatCurrentDocument()
		/// <summary>
		///  Formats the whole document
		/// </summary>
		[Command("Text Editor::Ctrl+Shift+d")]
		public void FormatCurrentDocument() {
			TextSelection sel = Selection;
			Point p = GetPoint();

			sel.StartOfDocument();
			sel.EndOfDocument(true);

			DTE.ExecuteCommand("Edit.FormatSelection");

			MoveToPoint(p);

		}
		#endregion
		#region public void ActOnTab()
		/// <summary>
		///  Acts on the tab key press
		/// </summary>
		[Command("Text Editor::Ctrl+<")]
		public void ActOnTab() {
			TextSelection sel = Selection;
			Point p1 = GetPoint(sel.TopPoint);
			Point p2 = GetPoint(sel.BottomPoint);

			if (!MoveToXmlDocTagSimple(true)) {
				if (p1.Equals(p2)) {
					sel.Insert("\t");
				} else {
					MoveToPoint(p1);
					MoveToPoint(p2, true);
					DTE.ExecuteCommand("Edit.IncreaseLineIndent");
				}
			}
			sel.ActivePoint.TryToShow();

		}
		#endregion
		#region public void ActOnShiftTab()
		/// <summary>
		///  Acts on Shift tab key press
		/// </summary>
		[Command("Text Editor::Ctrl+Shift+>")]
		public void ActOnShiftTab() {
			TextSelection sel = Selection;
			Point p1 = GetPoint(sel.TopPoint);
			Point p2 = GetPoint(sel.BottomPoint);

			if (!MoveToXmlDocTagSimple(false)) {
				if (p1.Equals(p2)) {
					Point p = GetPoint();
					sel.CharLeft(true);
					if (Regex.IsMatch(sel.Text, "^[\\t ]$")) {
						sel.Delete();
					} else {
						MoveToPoint(p);
					}
				} else {
					MoveToPoint(p1);
					MoveToPoint(p2, true);
					DTE.ExecuteCommand("Edit.DecreaseLineIndent");
				}
			}
			sel.ActivePoint.TryToShow();

		}
		#endregion
		/* *******************************************************************
		 *  private methods
		 * *******************************************************************/
		#region private bool MoveToXmlDocTagSimple(bool forward = true)
		/// <summary>
		///  Moves the cursor to the next/previos xml doc comment tag
		/// </summary>
		/// <param name="forward">Should we go to the next, set to true. Go to previous, set to false.</param>
		/// <returns>True if we are in an XML doc tag, false otherwize</returns>
		private bool MoveToXmlDocTagSimple(bool forward = true) {
			Point p = GetPoint();
			TextSelection sel = (TextSelection)DTE.ActiveWindow.Document.Selection;
			sel.StartOfLine();

			string documentationPattern = "///";
			if (IsBasic) {
				documentationPattern = "'''";
			}

			if (sel.FindPattern("^[ \\t]*" + documentationPattern, (int)vsFindOptions.vsFindOptionsRegularExpression)) {
				sel.Collapse();
				Point p2 = GetPoint();
				//if we is positioned before the /// or if /// is on another line, skip it
				if (p.Y != p2.Y | p.X <= p2.X - 3) {
					MoveToPoint(p);
					return false;
				}
				MoveToPoint(p);
				if (forward) {
					if (sel.FindPattern("\\<[^/\\>]+\\>", (int)vsFindOptions.vsFindOptionsRegularExpression)) {
						if (sel.ActivePoint.EqualTo(sel.TopPoint)) {
							sel.SwapAnchor();
						}
						sel.Collapse();
						p = GetPoint();
						//if the start tag is alone at the end of the line, try to find the next line
						sel.EndOfLine(true);
						if (Regex.IsMatch(sel.Text, "^\\s*$")) {
							if (sel.FindPattern(documentationPattern + " *", (int)vsFindOptions.vsFindOptionsRegularExpression)) {
								sel.Collapse();
								return true;
							}
						}
						MoveToPoint(p);
						return true;
					}
				} else {
					//skip current tag
					sel.FindPattern("\\<\\/[^/\\>]+\\>", (int)(vsFindOptions.vsFindOptionsRegularExpression + (int)vsFindOptions.vsFindOptionsBackwards));
					sel.Collapse();

					if (sel.FindPattern("\\<[^/\\>]+\\>", (int)(vsFindOptions.vsFindOptionsRegularExpression + (int)vsFindOptions.vsFindOptionsBackwards))) {
						if (sel.ActivePoint.EqualTo(sel.TopPoint)) {
							sel.SwapAnchor();
						}
						sel.Collapse();
						p = GetPoint();
						//if the start tag is alone at the end of the line, try to find the next line
						sel.EndOfLine(true);
						if (Regex.IsMatch(sel.Text, "^\\s*$")) {
							if (sel.FindPattern(documentationPattern + " *", (int)vsFindOptions.vsFindOptionsRegularExpression)) {
								sel.Collapse();
								return true;
							}
						}
						MoveToPoint(p);
						return true;
					}

				}
				MoveToPoint(p);
				return true;
			}
			MoveToPoint(p);
			return false;

		}
		#endregion
		#region private string AssertEndElement(string text, string elem, string[] parentElements)
		/// <summary>
		/// Makes sure that the found elements with the supplied name has an end tag
		/// </summary>
		/// <param name="text">The text to parse</param>
		/// <param name="elem">The name of the element to find</param>
		/// <param name="parentElements">The list of allowed parent elements</param>
		/// <returns>The resulting text</returns>
		/// <remarks>Is designed to add missing &lt;/li&gt; tags.</remarks>
		private string AssertEndElement(string text, string elem, string[] parentElements) {
			string newText = text;
			int offset = 0;
			string pattern = "<\\/?(?:" + elem;
			foreach (string pe in parentElements) {
				pattern += "|" + pe;
			}
			pattern += ")\\b[^>]*>";
			MatchCollection matches = Regex.Matches(text, pattern, RegexOptions.IgnoreCase);
			Match startMatch = default(Match);
			foreach (Match m in matches) {
				//is this a start tag?

				if (m.Groups[0].Value.StartsWith("</" + elem)) {
					//do we have a end tag without an start tag?

					if (startMatch == null) {
						//remove end tag
						newText = newText.Substring(0, m.Index + offset) + newText.Substring(m.Index + offset + m.Length);
						offset -= m.Length;
					}
					startMatch = null;
					continue;
				}
				if ((startMatch != null)) {
					string endTag = "</" + elem + ">";
					newText = newText.Substring(0, m.Index + offset) + endTag + newText.Substring(m.Index + offset);
					offset += endTag.Length;
				}

				if (m.Groups[0].Value.StartsWith("<" + elem) & !m.Groups[0].Value.EndsWith("/>")) {
					startMatch = m;
				}
			}
			return newText;
		}
		#endregion
		#region private string EnsureAttribute(string text, string tagName, string attribute, string value, string valueFromAttribute = null)
		/// <summary>
		/// Searches for tags and adds an attribute if they don't have them
		/// </summary>
		/// <param name="text">The text to search in</param>
		/// <param name="tagName">The name of the tag to search for</param>
		/// <param name="attribute">The name of the attribute to ensure</param>
		/// <param name="value">The value of the attribute to use if no attribute is found</param>
		/// <param name="valueFromAttribute"></param>
		/// <returns>The replaced text</returns>
		private string EnsureAttribute(string text, string tagName, string attribute, string value, string valueFromAttribute = null) {
			MatchCollection ms = Regex.Matches(text, "<" + tagName + "([^>]*)>", RegexOptions.IgnoreCase);
			int textOffset = 0;

			foreach (Match m in ms) {
				var attr = ExtractAttributes(m.Groups[1].Value);
				if (!attr.ContainsKey(attribute.ToLower())) {
					if ((valueFromAttribute != null)) {
						if (attr.ContainsKey(valueFromAttribute.ToLower())) {
							value = attr[valueFromAttribute.ToLower()];
						}
					}
					attr.Add(attribute.ToLower(), value);
					string tag = BuildTag(tagName, attr);
					text = text.Substring(0, m.Index + textOffset) + tag + text.Substring(m.Index + m.Length + textOffset);
					textOffset += tag.Length - m.Length;

				}
			}
			return text;
		}
		#endregion
		#region private string BuildTag(string tag, Dictionary<string,string> attr)
		/// <summary>
		/// Builds a start tag from a name and a set of attributes
		/// </summary>
		/// <param name="tag">The name of the tag</param>
		/// <param name="attr">The attributes of the tag</param>
		/// <returns>The built tag</returns>
		private string BuildTag(string tag, Dictionary<string, string> attr) {
			string t = null;
			t = "<" + tag;
			foreach (string key in attr.Keys) {
				string value = Convert.ToString(attr[key]);
				if (value == null) {
					value = string.Empty;
				}
				value = value.Replace("\"", "&quote;");
				t += " " + key.ToLower() + "=\"" + value + "\"";
			}
			t += ">";
			return t;
		}
		#endregion
		#region private string ExtractAttributesMatchEvaluator(Match m)
		/// <summary>
		/// Extracts attributes from the text and adds the found key/value pair to the ExtractAttributesList
		/// </summary>
		/// <param name="m">The <see cref="Match"/> containing the attribute.</param>
		/// <returns>An empty string if the match contained an attribute.</returns>
		private string ExtractAttributesMatchEvaluator(Match m) {
			string key = null;
			string value = null;
			if (m.Groups.Count == 3) {
				key = m.Groups[1].Value;
				value = m.Groups[2].Value;
			} else if (m.Groups.Count == 2) {
				key = m.Groups[1].Value;
				value = key;
			} else {
				return m.Groups[0].Value;
			}
			ExtractAttributesList.Add(key.ToLower(), value);
			return "";
		}
		#endregion
		#region private Dictionary<string,string> ExtractAttributes(string tag)
		/// <summary>
		/// Finds all attribute key/value pairs in the supplied tag.
		/// </summary>
		/// <param name="tag">The complete tag to parse</param>
		/// <returns>A <see cref="Hashtable"/> containing the found attributes</returns>
		private Dictionary<string, string> ExtractAttributes(string tag) {
			if (ExtractAttributesList == null) {
				ExtractAttributesList = new Dictionary<string, string>();
			} else {
				ExtractAttributesList.Clear();
			}

			//attributes with quotes
			tag = Regex.Replace(tag, "(\\S+)\\s*=\\s*\"([^\"]*)\"", ExtractAttributesMatchEvaluator, RegexOptions.Singleline);
			//attributes with single quotes
			tag = Regex.Replace(tag, "(\\S+)\\s*=\\s*'([^']*)'", ExtractAttributesMatchEvaluator, RegexOptions.Singleline);
			//attributes without quotes
			tag = Regex.Replace(tag, "(\\S+)\\s*=\\s*(\\S+)", ExtractAttributesMatchEvaluator, RegexOptions.Singleline);
			//attributes without value
			tag = Regex.Replace(tag, "(\\w\\S*)", ExtractAttributesMatchEvaluator, RegexOptions.Singleline);

			return ExtractAttributesList;
		}
		#endregion
		#region private string FixAttributesMatchEvaluator(Match m)
		/// <summary>
		/// When this is called on an attribute match where the name and value are divided in two matching groups
		/// this will replace the attribute with a lowercase name and quote the value if not already quoted.
		/// </summary>
		/// <param name="m">The Match containing the attribute</param>
		/// <returns>The match replacement</returns>
		/// <exception cref="ArgumentException">If can't process input, should only be two groups.</exception>
		private string FixAttributesMatchEvaluator(Match m) {
			if (m.Groups.Count == 1) {
				return m.Groups[0].Value.ToLower();
			}
			if (m.Groups.Count == 3) {
				string s = "";
				string orgs = m.Groups[0].Value;
				int startIndex = m.Groups[0].Index;
				int textOffset = 0;

				for (int i = 1; i <= m.Groups.Count - 1; i++) {
					for (int j = 0; j <= m.Groups[i].Captures.Count - 1; j++) {
						if ((s.Length - textOffset) < (m.Groups[i].Captures[j].Index - startIndex)) {
							//copy text to the group
							s += orgs.Substring(s.Length - textOffset, m.Groups[i].Captures[j].Index - startIndex - (s.Length - textOffset));
						}
						if (i == 1) {
							s += m.Groups[i].Captures[j].Value.ToLower();
						} else {
							KeyValuePair<string, string> kv = ParseAttribute(m.Groups[i].Captures[j].Value);
							string value = kv.Value;
							value = value.Replace("\"", "&quote;");
							string newText = kv.Key.ToLower() + "=\"" + value + "\"";
							s += newText;
							textOffset += newText.Length - m.Groups[i].Captures[j].Length;
						}

					}

				}
				if ((s.Length - textOffset) < orgs.Length) {
					s += orgs.Substring((s.Length - textOffset));
				}
				return s;
			}
			throw new ArgumentException("Can't process input, should only be two groups");
		}
		#endregion
		#region private KeyValuePair<string, string> ParseAttribute(string att)
		/// <summary>
		/// Parses an attribute key/value pair
		/// </summary>
		/// <param name="att">The text to parse</param>
		/// <returns>The found <see cref="KeyValuePair{string,string}"/> or nothing if no attribute was found</returns>
		/// <exception cref="ArgumentException">If more than one attribute was found.</exception>
		private KeyValuePair<string, string> ParseAttribute(string att) {
			KeyValuePair<string, string> kv = default(KeyValuePair<string, string>);
			var ht = ExtractAttributes(att);
			if (ht.Keys.Count == 0) {
				return kv;
			}
			if (ht.Keys.Count > 1) {
				throw new ArgumentException("more than one attribute was found");
			}
			IEnumerator enu = (ht.Keys.GetEnumerator());
			enu.MoveNext();
			string key = Convert.ToString(enu.Current);
			string value = Convert.ToString(ht[key]);
			kv = new KeyValuePair<string, string>(key, value);
			return kv;
		}
		#endregion
	}
}