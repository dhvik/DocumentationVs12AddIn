
using System;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using EnvDTE;

namespace DocumentationVs12AddIn.Commands {
	/// <summary>
	/// Summary description for PasteTemplateCommand.
	/// </summary>
	/// <remarks>
	/// 2013-04-09 dan: Created
	/// </remarks>
	public class PasteTemplateCommand : CommandBase {
		#region public void PasteTemplate()
		/// <summary>
		/// Creates a template insertion by the written keyword
		/// </summary>
		/// <remarks>
		/// Currently the following keywords is defined
		///  fori writes a <c><![CDATA[for(int i=0;i<|.Count;i++) {...]]></c> template
		///  tryc writes a try catch template
		///  trycf writes a try catch finally template
		///  tryf writes a try finally template
		/// </remarks>
		[Command("Text Editor::Ctrl+Shift+j")]
		public void PasteTemplate() {
			if ((DTE.ActiveWindow.Type != vsWindowType.vsWindowTypeDocument)) {
				return;
			}
			if ((!IsCSharp & !IsBasic)) {
				return;
			}

			TextSelection sel = Selection;

			//get templateName
			string templateName = CutLeftWord();
			string template = null;

			//store startpoint
			Point startPoint = GetPoint();

			//get indent level
			sel.SelectLine();
			sel.SwapAnchor();
			string indent = sel.Text;
			indent = Regex.Replace(indent, "^([ \\t]*).*", "$1", RegexOptions.Singleline);

			//move to insertionpoint
			MoveToPoint(startPoint);

			Match m;
			switch (templateName) {
				case "fori":
					template = "for(int i=0;i<|;i++) {" + Environment.NewLine + indent + "\t" + Environment.NewLine + indent + "}";
					break;
				case "forj":
					template = "for(int j=0;j<|;j++) {" + Environment.NewLine + indent + "\t" + Environment.NewLine + indent + "}";
					break;
				case "fork":
					template = "for(int k=0;k<|;k++) {" + Environment.NewLine + indent + "\t" + Environment.NewLine + indent + "}";
					break;
				case "tryc":
					template = "try {" + Environment.NewLine + indent + "\t" + "|" + Environment.NewLine + indent + "} catch(Exception ex) {" + Environment.NewLine + indent + "\t" + Environment.NewLine + indent + "}";
					break;
				case "trycm":
					template = "try {" + Environment.NewLine + indent + "\t" + "|" + Environment.NewLine + indent + "} catch {}";
					break;
				case "trycf":
					template = "try {" + Environment.NewLine + indent + "\t" + "|" + Environment.NewLine + indent + "} catch(Exception ex) {" + Environment.NewLine + indent + "\t" + Environment.NewLine + indent + "} finally {" + Environment.NewLine + indent + "\t" + Environment.NewLine + indent + "}";
					break;
				case "tryf":
					template = "try {" + Environment.NewLine + indent + "\t" + "|" + Environment.NewLine + indent + "} finally {" + Environment.NewLine + indent + "\t" + Environment.NewLine + indent + "}";
					break;
				case "test":
					template = "[Test]" + Environment.NewLine + indent + "public void |() {" + Environment.NewLine + indent + "\t" + Environment.NewLine + indent + "}" + Environment.NewLine;
					break;
				case "gp":
					template = "{" + Environment.NewLine;
					template += indent + "\t" + "get {" + Environment.NewLine;
					template += indent + "\t" + "\t" + "|" + Environment.NewLine;
					template += indent + "\t" + "}" + Environment.NewLine;
					template += indent + "}" + Environment.NewLine;
					break;
				case "sp":
					template = "{" + Environment.NewLine;
					template += indent + "\t" + "set {" + Environment.NewLine;
					template += indent + "\t" + "\t" + "|" + Environment.NewLine;
					template += indent + "\t" + "}" + Environment.NewLine;
					template += indent + "}" + Environment.NewLine;
					break;
				case "gsp":
					template = "{" + Environment.NewLine;
					template += indent + "\t" + "get {" + Environment.NewLine;
					template += indent + "\t" + "\t" + "|" + Environment.NewLine;
					template += indent + "\t" + "}" + Environment.NewLine;
					template += indent + "\t" + "set {" + Environment.NewLine;
					template += indent + "\t" + "\t" + "" + Environment.NewLine;
					template += indent + "\t" + "}" + Environment.NewLine;
					template += indent + "}" + Environment.NewLine;
					break;
				case "disp":
					sel.SelectLine();
					sel.SwapAnchor();

					m = Regex.Match(sel.Text, "(?<Variable>\\S+)\\s*$", RegexOptions.Singleline);
					if (m.Success) {
						string var = m.Groups["Variable"].Value;

						template = "";
						template += indent + "if(" + var + " != null) {" + Environment.NewLine;
						template += indent + "\t" + var + ".Dispose();" + Environment.NewLine;
						template += indent + "\t" + var + " = null;" + Environment.NewLine;
						template += indent + "}|" + Environment.NewLine;
					}
					break;
				case ";":
					sel.SelectLine();
					sel.SwapAnchor();
					//event declaration
					//ex
					//public event EventHandler MyEvent
					m = Regex.Match(sel.Text, "(?<AccessModifier>(?:(?:private|public|internal|protected|static)\\s+)*)\\bevent\\s+(?<Type>[^\\s]+Handler(?:\\s*\\<[^\\>]+\\>)?)\\s+(?<Name>[^\\s=;]+)\\s*$", RegexOptions.Singleline);
					if (m.Success) {
						string type = m.Groups["Type"].Value.Trim();
						string argsType = Regex.Replace(type, "Handler", "Args");
						string name = m.Groups["Name"].Value.Trim();



						string descriptor = null;
						// accessModifier
						if (m.Groups["AccessModifier"].Success) {
							descriptor += m.Groups["AccessModifier"].Value.Trim() + " ";
						} else {
							descriptor += "public ";
						}
						descriptor += "event ";
						descriptor += type + " " + name;

						m = Regex.Match(type, "EventHandler\\s*\\<\\s*(?<ArgsType>[^\\>]+)\\s*\\>", RegexOptions.Singleline);
						if (m.Success) {
							argsType = m.Groups["ArgsType"].Value.Trim();
						} else if (type.IndexOfAny(new char[] {
					',',
					'<',
					'>'
				}) > -1) {
							MessageBox.Show("The eventhandler type " + type + " is not supported");
							return;
						}
						string extraDescription = "";
						m = Regex.Match(name, "^(\\S+)Changed$");
						if (m.Success) {
							extraDescription = "the " + m.Groups[1].Value + " property is changed.";
						}
						sel.Delete();
						template = "";

						template += indent + "#region " + descriptor + Environment.NewLine;
						template += indent + "/// <summary>" + Environment.NewLine;
						template += indent + "/// This event is fired when " + extraDescription + "|" + Environment.NewLine;
						template += indent + "/// </summary>" + Environment.NewLine;
						template += indent + descriptor + ";" + Environment.NewLine;
						template += indent + "#endregion" + Environment.NewLine;

						descriptor = "protected virtual void On" + name + "(" + argsType + " e)";

						template += indent + "#region " + descriptor + Environment.NewLine;
						template += indent + "/// <summary>" + Environment.NewLine;
						template += indent + "/// Notifies the listeners of the " + name + " event" + Environment.NewLine;
						template += indent + "/// </summary>" + Environment.NewLine;
						template += indent + "/// <param name=\"e\">The argument to send to the listeners</param>" + Environment.NewLine;
						template += indent + descriptor + " {" + Environment.NewLine;
						template += indent + "\t" + "if(" + name + " != null) {" + Environment.NewLine;
						template += indent + "\t" + "\t" + name + "(this,e);" + Environment.NewLine;
						template += indent + "\t" + "}" + Environment.NewLine;
						template += indent + "}" + Environment.NewLine;
						template += indent + "#endregion" + Environment.NewLine;

						break; // TODO: might not be correct. Was : Exit Select
					}

					//public static string Instance(=(null|.+))?
					m = Regex.Match(sel.Text, "(?<AccessModifier>(?:(?:private|public|internal|protected)\\s+)*)static\\s+(?<Type>[^\\s]+)\\s+Instance\\s*(?<Assignment>\\s*=\\s*.+)?\\s*$");
					if ((m.Success)) {
						sel.Delete();

						string assignment = m.Groups["Assignment"].Value.Trim();
						string type = m.Groups["Type"].Value.Trim();
						bool singletonPattern = false;

						string declaration = "";

						//Property start, accessModifier
						if (m.Groups["AccessModifier"].Success) {
							declaration += m.Groups["AccessModifier"].Value.Trim() + " ";
						} else {
							declaration += "public ";
						}

						//Property type and name
						declaration += "static " + type + " Instance";

						template = indent + "#region " + declaration + Environment.NewLine;
						template += indent + "/// <summary>" + Environment.NewLine;
						template += indent + "/// Returns the singleton instance of the " + GetSeeXmlDocTag(type) + " class." + Environment.NewLine;
						template += indent + "/// </summary>" + Environment.NewLine;
						template += indent + declaration + " {" + Environment.NewLine;

						//Get method
						template += indent + "\t" + "get {";
						if (assignment.Contains("null") || assignment.Length == 0) {
							template += Environment.NewLine;
							template += indent + "\t" + "\t" + "if(_instance == null) {" + Environment.NewLine;
							template += indent + "\t" + "\t" + "\t" + "lock(InstanceLock) {" + Environment.NewLine;
							template += indent + "\t" + "\t" + "\t" + "\t" + "if(_instance == null) {" + Environment.NewLine;
							template += indent + "\t" + "\t" + "\t" + "\t" + "\t" + "_instance = new " + type + "(|);" + Environment.NewLine;
							template += indent + "\t" + "\t" + "\t" + "\t" + "}" + Environment.NewLine;
							template += indent + "\t" + "\t" + "\t" + "}" + Environment.NewLine;
							template += indent + "\t" + "\t" + "}" + Environment.NewLine;
							template += indent + "\t" + "\t" + "return _instance;" + Environment.NewLine;
							template += indent + "\t";
							singletonPattern = true;
						} else {
							template += " return _instance; ";
						}
						template += "}" + Environment.NewLine;
						template += indent + "}" + Environment.NewLine;
						//private variable
						template += indent + "private " + (singletonPattern ? "volatile " : "readonly ") + "static " + type + " _instance";
						if (assignment.Length > 0 && !singletonPattern) {
							template += " " + assignment;
						}
						template += ";" + Environment.NewLine;
						if (singletonPattern)
							template += indent + "private static readonly object InstanceLock = new object();" + Environment.NewLine;
						template += indent + "#endregion" + Environment.NewLine;
						break; // TODO: might not be correct. Was : Exit Select
					}

					//field declaration, convert to property
					//ex
					//public string MyString(=(null|.+))?
					//(?:
					//	\<				    #start of attribute marker
					//	(?>				    #greedy sub
					//		(?<T>		    #Define terms for attributes
					//			(?:		    #non capturing group
					//
					//				[^\<\>]	#any character that not is <>
					//			)+			#more than one
					//		)			    #end of term definition
					//			|		    #or we get a
					//		\< (?<D>)	    #increase in the nested depth
					//			|		    #or we get a
					//		\>(?<-D>)	    #decrease in the nested depth
					//	)*				    #Any number of terms
					//	(?(D)(?!))		    #match depth
					//	\>				    #end of attribute marker
					//)?					    #parameters are optional

					//(?:\<(?>(?<T>(?:[^\<\>])+)|\<(?<D>)|\>(?<-D>))*(?(D)(?!))\>)?
					//2009-02-27, Dan Händevik - Now supports virtual, new and override keywords
					//2010-05-23, Dan Händevik, Added support for creating autoproperty when no assignment is made.
					m = Regex.Match(sel.Text, "(?<AccessModifier>(?:(?:private|public|internal|protected)\\s+)*)(?<New>new\\s)?\\s*(?<Override>override\\s)?\\s*(?<Virtual>virtual\\s)?\\s*(?<Static>static\\s)?\\s*(?<Type>[^\\s]+(?:\\<(?>(?<T>(?:[^\\<\\>])+)|\\<(?<D>)|\\>(?<-D>))*(?(D)(?!))\\>)?)\\s+(?<Name>[^\\s=;]+)(?<Assignment>\\s*=\\s*.+)?\\s*$");
					if ((m.Success)) {
						sel.Delete();
						template = indent;

						string assignment = m.Groups["Assignment"].Value.Trim();
						string accessModifier = "public";
						if (m.Groups["AccessModifier"].Success) {
							accessModifier = m.Groups["AccessModifier"].Value.Trim();
						}
						string type = m.Groups["Type"].Value.Trim();
						string name = m.Groups["Name"].Value.Trim();
						if (name.StartsWith("_") & name.Length > 1) {
							name = name.Substring(1, 1).ToUpper() + name.Substring(2);
							accessModifier = "public";
						}
						string localName = "_" + name.ToLower()[0] + name.Substring(1);
						bool IsStatic = (m.Groups["Static"].Value.Length > 0);
						bool IsVirtual = (m.Groups["Virtual"].Value.Length > 0);
						bool IsOverride = (m.Groups["Override"].Value.Length > 0);
						bool IsNew = (m.Groups["New"].Value.Length > 0);
						bool HasAssignment = (assignment.Length > 0 && !assignment.Contains("null"));
						bool autoProperty = !HasAssignment;

						//Property start, accessModifier
						template += accessModifier + " ";

						if (IsStatic) {
							template += "static ";
						}
						if (IsVirtual) {
							template += "virtual ";
						}
						if (IsOverride) {
							template += "override ";
						}
						if (IsNew) {
							template += "new ";
						}
						//Property type and name
						template += type + " " + name + " {";


						//Get method
						if (HasAssignment) {
							template += Environment.NewLine;
							template += indent + "\t" + "get {return " + localName + ";}" + Environment.NewLine;
						} else {
							template += "get;";
						}

						//set method
						if (HasAssignment) {
							template += indent + "\t" + "set {" + localName + " = value;}" + Environment.NewLine;
							template += indent + "}" + Environment.NewLine;

							//private variable
							template += indent + "private ";
							if (IsStatic) {
								template += "static ";
							}
							template += type + " " + localName;
							if (assignment.Length > 0) {
								template += " " + assignment;
							}
							template += ";" + Environment.NewLine;
						} else {
							template += "set;}" + Environment.NewLine;
						}
						break; // TODO: might not be correct. Was : Exit Select
					}
					MoveToPoint(startPoint);
					template = ";";
					break;
				case "}":
					sel.SelectLine();
					sel.SwapAnchor();
					//set {var=value;}
					m = Regex.Match(sel.Text, "set\\s*{\\s*(?<Name>[^\\s=;]+)\\s*=\\s*value\\s*;\\s*$", RegexOptions.Singleline);
					if (m.Success) {
						sel.Delete();
						template = "";
						template += indent + "set {" + Environment.NewLine;
						template += indent + "\t" + "if(" + m.Groups["Name"].Value + " != value) {" + Environment.NewLine;
						template += indent + "\t" + "\t" + m.Groups["Name"].Value + " = value;" + Environment.NewLine;
						template += indent + "\t" + "\t" + "|" + Environment.NewLine;
						template += indent + "\t" + "}" + Environment.NewLine;
						template += indent + "}" + Environment.NewLine;
						break; // TODO: might not be correct. Was : Exit Select
					}
					//get {return var;}
					//2010-05-23, fixed replacement to single line assignment
					m = Regex.Match(sel.Text, "get\\s*{\\s*return (?<Name>[^\\s=;]+)\\s*;\\s*$", RegexOptions.Singleline);
					if (m.Success) {
						sel.Delete();
						string name = m.Groups["Name"].Value;
						template = "";
						template += indent + "get{ return " + name + " ?? (" + name + " = |); }" + Environment.NewLine;
						//template += indent + "get {" + vbCrLf
						//template += indent + vbTab + "if(" + name + " == null) {" + vbCrLf
						//template += indent + vbTab + vbTab + name + " = |;" + vbCrLf
						//template += indent + vbTab + "}" + vbCrLf
						//template += indent + vbTab + "return " + name + ";" + vbCrLf
						//template += indent + "}" + vbCrLf
						break; // TODO: might not be correct. Was : Exit Select
					}
					//auto property declaration, convert to backing field
					//ex
					//public string MyString {get;?set;?
					//(?:
					//	\<				    #start of attribute marker
					//	(?>				    #greedy sub
					//		(?<T>		    #Define terms for attributes
					//			(?:		    #non capturing group
					//
					//				[^\<\>]	#any character that not is <>
					//			)+			#more than one
					//		)			    #end of term definition
					//			|		    #or we get a
					//		\< (?<D>)	    #increase in the nested depth
					//			|		    #or we get a
					//		\>(?<-D>)	    #decrease in the nested depth
					//	)*				    #Any number of terms
					//	(?(D)(?!))		    #match depth
					//	\>				    #end of attribute marker
					//)?					    #parameters are optional

					//(?:\<(?>(?<T>(?:[^\<\>])+)|\<(?<D>)|\>(?<-D>))*(?(D)(?!))\>)?
					//2009-02-27, Dan Händevik - Now supports virtual, new and override keywords
					//2010-05-23, Dan Händevik - Added mapping from auto property to backing field
					m = Regex.Match(sel.Text, "(?<AccessModifier>(?:(?:private|public|internal|protected)\\s+)*)(?<New>new\\s)?\\s*(?<Override>override\\s)?\\s*(?<Virtual>virtual\\s)?\\s*(?<Static>static\\s)?\\s*(?<Type>[^\\s]+(?:\\<(?>(?<T>(?:[^\\<\\>])+)|\\<(?<D>)|\\>(?<-D>))*(?(D)(?!))\\>)?)\\s+(?<Name>[^\\s=;]+)\\s*{\\s*(?:(?<get>get;\\s*)|(?<set>set;\\s*))*\\s*$", RegexOptions.Singleline);
					if ((m.Success)) {
						sel.Delete();
						template = indent;

						string accessModifier = "public";
						if (m.Groups["AccessModifier"].Success) {
							accessModifier = m.Groups["AccessModifier"].Value.Trim();
						}
						string type = m.Groups["Type"].Value.Trim();
						string name = m.Groups["Name"].Value.Trim();
						if (name.StartsWith("_") & name.Length > 1) {
							name = name.Substring(1, 1).ToUpper() + name.Substring(2);
							accessModifier = "public";
						}
						string localName = "_" + name.ToLower()[0] + name.Substring(1);
						bool IsStatic = (m.Groups["Static"].Value.Length > 0);
						bool IsVirtual = (m.Groups["Virtual"].Value.Length > 0);
						bool IsOverride = (m.Groups["Override"].Value.Length > 0);
						bool IsNew = (m.Groups["New"].Value.Length > 0);
						bool HasSetter = (m.Groups["set"].Value.Length > 0);
						bool HasGetter = (m.Groups["get"].Value.Length > 0);

						//Property start, accessModifier
						template += accessModifier + " ";

						if (IsStatic) {
							template += "static ";
						}
						if (IsVirtual) {
							template += "virtual ";
						}
						if (IsOverride) {
							template += "override ";
						}
						if (IsNew) {
							template += "new ";
						}
						//Property type and name
						template += type + " " + name + " {" + Environment.NewLine;

						//Get method
						if (HasGetter) {
							template += indent + "\t" + "get {";
							template += " return " + localName + "; ";
							template += "}" + Environment.NewLine;
						}

						//set method
						if (HasSetter) {
							template += indent + "\t" + "set {";
							template += " " + localName + " = value; ";
							template += "}" + Environment.NewLine;
						}
						template += indent + "}" + Environment.NewLine;

						//private variable
						template += indent + "private ";
						if (IsStatic) {
							template += "static ";
						}
						template += type + " " + localName;
						template += ";" + Environment.NewLine;
						break; // TODO: might not be correct. Was : Exit Select
					}

					MoveToPoint(startPoint);
					template = "}";

					break; // TODO: might not be correct. Was : Exit Select

					break;
				case "{":
					sel.SelectLine();
					sel.SwapAnchor();
					//convert foreach to fori
					m = Regex.Match(sel.Text, "foreach\\s*\\(\\s*(?<Type>.+)\\s+(?<Variable>[^\\s]+)\\s+in\\s+(?<Collection>[^\\s]+)\\s*\\)\\s*$", RegexOptions.Singleline);
					if (m.Success) {
						sel.Delete();
						template = "";
						template += indent + "for(int i=0;i<" + m.Groups["Collection"].Value + ".Count;i++) {" + Environment.NewLine;
						template += indent + "\t" + m.Groups["Type"].Value + " " + m.Groups["Variable"].Value + " = " + m.Groups["Collection"].Value + "[i];" + Environment.NewLine;

						break; // TODO: might not be correct. Was : Exit Select
					}

					MoveToPoint(startPoint);
					template = "{";

					break;
				default:
					template = templateName;
					break;
			}

			//insert the template
			sel.Insert(template);

			//find marker
			//if a marker exists in the teplate
			if (template != null && template.Contains("|")) {
				if (sel.FindPattern("|", (int)vsFindOptions.vsFindOptionsBackwards)) {
					sel.Delete();
				}
			}


		}
		#endregion
	}
}