using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using DocumentationVs12AddIn.Commands;
using EnvDTE;
using EnvDTE80;
using Extensibility;
using log4net;
using log4net.Config;

namespace DocumentationVs12AddIn {
	/// <summary>The object for implementing an Add-in.</summary>
	/// <seealso class='IDTExtensibility2' />
	public sealed class Connect : IDTExtensibility2, IDTCommandTarget {

		private readonly ILog Log = LogManager.GetLogger(typeof(Connect));

		private Dictionary<string, CommandInfo> _registredCommands;
		private DTE2 _applicationObject;
		private AddIn _addInInstance;

		/* *******************************************************************
		 *  Constructors
		 * *******************************************************************/
		#region public Connect()
		/// <summary>
		/// Initializes a new instance of the <b>Connect</b> class.
		/// </summary>
		public Connect() {
			//Connect logging...
			var assemblyFile = new FileInfo(GetType().Assembly.Location);
			if (assemblyFile.Exists && assemblyFile.Directory != null) {
				var configFile = new FileInfo(Path.Combine(assemblyFile.Directory.FullName, "log4net.config"));
				if (configFile.Exists)
					XmlConfigurator.Configure(configFile);
			}
		}
		#endregion
		/* *******************************************************************
		 *  Methods
		 * *******************************************************************/
		#region public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
		/// <summary>Implements the OnConnection method of the IDTExtensibility2 interface. Receives notification that the Add-in is being loaded.</summary>
		/// <param name="application">Root object of the host application.</param>
		/// <param name="connectMode">Describes how the Add-in is being loaded.</param>
		/// <param name="addInInst">Object representing this Add-in.</param>
		/// <param name="custom"></param>
		public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom) {

			Log.Debug("mode:" + connectMode);
			try {
				_applicationObject = (DTE2)application;
				_addInInstance = (AddIn)addInInst;

				var commands2 = (Commands2)_applicationObject.Commands;
				var commandInfos = CalculateCommandInfos();

				_registredCommands = commandInfos.ToDictionary(x => x.Attribute.CommandName, x => x);

				//add commands to ui on setup only (This is permanent unless you run 'devenv.exe /resetaddin DocumentationVs12AddIn.Connect')
				if (connectMode == ext_ConnectMode.ext_cm_UISetup) {
					var contextGuids = new object[] { };

					foreach (var commandInfo in commandInfos) {
						var dteCommand = commands2.AddNamedCommand2(_addInInstance, commandInfo.Attribute.Name, commandInfo.Attribute.Name, commandInfo.Attribute.Description,
																	true, commandInfo.Attribute.IconId, ref contextGuids);
						if (!string.IsNullOrEmpty(commandInfo.Attribute.KeyBinding))
							BindPreserve(dteCommand.Name, commandInfo.Attribute.KeyBinding);
					}

				}


			} catch (Exception e) {
				MessageBox.Show(e.Message + Environment.NewLine + e.StackTrace, "Error loading addin:" + GetType().Assembly.FullName);
			}
		}
		#endregion
		#region private List<CommandInfo> CalculateCommandInfos()
		/// <summary>
		/// Calculates the CommandInfos in the project.
		/// </summary>
		/// <remarks>These are based on found classes inheriting from CommandBase and methods decorated with the <see cref="CommandAttribute"/> attribute</remarks>
		/// <returns></returns>
		private List<CommandInfo> CalculateCommandInfos() {
			var commandInfos = new List<CommandInfo>();
			var commandTypes =
				GetType().Assembly.GetTypes().Where(t => typeof(CommandBase).IsAssignableFrom(t) && !t.IsAbstract).ToList();

			//iterate all command types and find the command methods
			foreach (var commandType in commandTypes) {
				var commandInstance = (CommandBase)Activator.CreateInstance(commandType);
				commandInstance.DTE = _applicationObject;
				var methods = commandType.GetMethods(BindingFlags.Instance | BindingFlags.Public);
				foreach (var methodInfo in methods) {
					if (methodInfo.GetParameters().Length != 0)
						continue;
					var commandAttribute =
						methodInfo.GetCustomAttributes(typeof(CommandAttribute), false).FirstOrDefault() as CommandAttribute;
					if (commandAttribute == null)
						continue;
					if (string.IsNullOrEmpty(commandAttribute.Name))
						commandAttribute.Name = methodInfo.Name;
					if (string.IsNullOrEmpty(commandAttribute.Description))
						commandAttribute.Description = commandAttribute.Name;
					commandAttribute.CommandName = _addInInstance.Name + ".Connect." + commandAttribute.Name;
					//create a command info for every command method
					var commandInfo = new CommandInfo(commandInstance, methodInfo, commandAttribute);
					commandInfos.Add(commandInfo);
				}
			}
			return commandInfos;
		}
		#endregion
		#region public void OnDisconnection(ext_DisconnectMode disconnectMode, ref Array custom)
		/// <summary>Implements the OnDisconnection method of the IDTExtensibility2 interface. Receives notification that the Add-in is being unloaded.</summary>
		/// <param name="disconnectMode">Describes how the Add-in is being unloaded.</param>
		/// <param name="custom">Array of parameters that are host application specific.</param>
		public void OnDisconnection(ext_DisconnectMode disconnectMode, ref Array custom) {
		}
		#endregion
		#region public void OnAddInsUpdate(ref Array custom)
		/// <summary>Implements the OnAddInsUpdate method of the IDTExtensibility2 interface. Receives notification when the collection of Add-ins has changed.</summary>
		/// <param name="custom">Array of parameters that are host application specific.</param>
		public void OnAddInsUpdate(ref Array custom) {
		}
		#endregion
		#region public void OnStartupComplete(ref Array custom)
		/// <summary>Implements the OnStartupComplete method of the IDTExtensibility2 interface. Receives notification that the host application has completed loading.</summary>
		/// <param name="custom">Array of parameters that are host application specific.</param>
		public void OnStartupComplete(ref Array custom) {
		}
		#endregion
		#region public void OnBeginShutdown(ref Array custom)
		/// <summary>Implements the OnBeginShutdown method of the IDTExtensibility2 interface. Receives notification that the host application is being unloaded.</summary>
		/// <param name="custom">Array of parameters that are host application specific.</param>
		public void OnBeginShutdown(ref Array custom) {
		}
		#endregion
		#region public void QueryStatus(string commandName, vsCommandStatusTextWanted neededText, ref vsCommandStatus status, ref object commandText)
		/// <summary>
		/// 
		/// </summary>
		/// <param name="commandName"></param>
		/// <param name="neededText"></param>
		/// <param name="status"></param>
		/// <param name="commandText"></param>
		public void QueryStatus(string commandName, vsCommandStatusTextWanted neededText, ref vsCommandStatus status, ref object commandText) {
			if (neededText != vsCommandStatusTextWanted.vsCommandStatusTextWantedNone)
				return;
			if (!_registredCommands.ContainsKey(commandName))
				return;
			var registredCommandInfo = _registredCommands[commandName];
			registredCommandInfo.Command.QueryStatus(neededText, ref status, ref commandText);
		}

		#endregion
		#region public void Exec(string commandName, vsCommandExecOption executeOption, ref object varIn, ref object varOut, ref bool handled)
		/// <summary>
		/// 
		/// </summary>
		/// <param name="commandName"></param>
		/// <param name="executeOption"></param>
		/// <param name="varIn"></param>
		/// <param name="varOut"></param>
		/// <param name="handled"></param>
		public void Exec(string commandName, vsCommandExecOption executeOption, ref object varIn, ref object varOut, ref bool handled) {
			if (handled)
				return;
			handled = false;
			if (executeOption != vsCommandExecOption.vsCommandExecOptionDoDefault)
				return;
			if (!_registredCommands.ContainsKey(commandName))
				return;
			var registredCommandInfo = _registredCommands[commandName];
			registredCommandInfo.Target.Invoke(registredCommandInfo.Command, null);
			//registredCommandInfo.Exec(executeOption, ref varIn, ref varOut);
			handled = true;
		}
		#endregion
		#region private void BindPreserve(string textcommand, string binding)
		/// <summary>
		/// Binds the supplied textcommand to the supplied binding
		/// </summary>
		/// <param name="textcommand"></param>
		/// <param name="binding"></param>
		public void BindPreserve(string textcommand, string binding) {
			var command = _applicationObject.Commands.Item(textcommand);
			var bindings = (object[])command.Bindings;
			var preserveLength = bindings.Length;
			var tmp = new object[preserveLength + 1];
			Array.Copy(bindings, tmp, bindings.Length);
			tmp[preserveLength] = binding;

			try {
				command.Bindings = tmp;
			} catch (Exception e) {
				MessageBox.Show("Error creating binding " + textcommand + " " + binding + " " + e.Message);
			}

		}
		#endregion
	}
}