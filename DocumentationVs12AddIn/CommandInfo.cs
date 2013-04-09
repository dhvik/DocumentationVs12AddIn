using System;
using System.Reflection;
using DocumentationVs12AddIn.Commands;
using EnvDTE;

namespace DocumentationVs12AddIn {
	public class CommandInfo {
		public CommandBase Command { get; private set; }


		public MethodInfo Target { get; private set; }
		public CommandAttribute Attribute { get; private set; }
		public CommandInfo(CommandBase commandBase, MethodInfo target, CommandAttribute commandAttribute) {
			Command=commandBase;
			Target=target;
			Attribute=commandAttribute;
		}
	}
}