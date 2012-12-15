using System;
using System.Windows.Forms;
using Extensibility;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.CommandBars;
using System.Resources;
using System.Reflection;
using System.Globalization;

namespace TomasAddin
{
	/// <summary>The object for implementing an Add-in.</summary>
	/// <seealso class='IDTExtensibility2' />
	public class Connect : IDTExtensibility2, IDTCommandTarget
	{
	    const string _addInNamespace = "TomasAddin.Connect.";

        const string _commandName_AttachToIIS = "AttachToIIS";
	    const string _commandName_SaveAndBuildExternal = "SaveAndBuildWithExternal";

	    const string _fullCommandName_SaveAndBuildExternal = _addInNamespace + _commandName_SaveAndBuildExternal;
	    const string _fullCommandName_AttachToIIS = _addInNamespace + _commandName_AttachToIIS;

	    /// <summary>Implements the constructor for the Add-in object. Place your initialization code within this method.</summary>
		public Connect()
		{
		}

		/// <summary>Implements the OnConnection method of the IDTExtensibility2 interface. Receives notification that the Add-in is being loaded.</summary>
		/// <param term='application'>Root object of the host application.</param>
		/// <param term='connectMode'>Describes how the Add-in is being loaded.</param>
		/// <param term='addInInst'>Object representing this Add-in.</param>
		/// <seealso class='IDTExtensibility2' />
		public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
		{
			_applicationObject = (DTE2)application;
			_addInInstance = (AddIn)addInInst;
			if(connectMode == ext_ConnectMode.ext_cm_UISetup)
			{
				var contextGUIDS = new object[] { };
				var commands = (Commands2)_applicationObject.Commands;
				const string toolsMenuName = "Tools";

				//Place the command on the tools menu.
				//Find the MenuBar command bar, which is the top-level command bar holding all the main menu items:
				var menuBarCommandBar = ((Microsoft.VisualStudio.CommandBars.CommandBars)_applicationObject.CommandBars)["MenuBar"];

				//Find the Tools command bar on the MenuBar command bar:
				var toolsControl = menuBarCommandBar.Controls[toolsMenuName];
				var toolsPopup = (CommandBarPopup)toolsControl;

				//This try/catch block can be duplicated if you wish to add multiple commands to be handled by your Add-in,
				//  just make sure you also update the QueryStatus/Exec method to include the new command names.
				try
				{
					//Add a command to the Commands collection:
				    var saveAndBuildParallel = commands.AddNamedCommand2(_addInInstance,
				                                                _commandName_SaveAndBuildExternal,
				                                                "Save & Build",
				                                                "Saves all files and builds the solution in parallel using the External Tools 7",
				                                                true,
				                                                59,
				                                                ref contextGUIDS,
				                                                (int) vsCommandStatus.vsCommandStatusSupported + (int) vsCommandStatus.vsCommandStatusEnabled,
				                                                (int) vsCommandStyle.vsCommandStylePictAndText,
				                                                vsCommandControlType.vsCommandControlTypeButton);
                    
                    var attachToIIS = commands.AddNamedCommand2(_addInInstance,
				                                                _commandName_AttachToIIS,
                                                                "Attach to IIS",
				                                                "Attaches to the current IIS",
				                                                true,
				                                                59,
				                                                ref contextGUIDS,
				                                                (int) vsCommandStatus.vsCommandStatusSupported +
				                                                (int) vsCommandStatus.vsCommandStatusEnabled,
				                                                (int) vsCommandStyle.vsCommandStylePictAndText,
				                                                vsCommandControlType.vsCommandControlTypeButton);

					//Add a control for the command to the tools menu:
					if((saveAndBuildParallel != null) && (toolsPopup != null))
					{
						saveAndBuildParallel.AddControl(toolsPopup.CommandBar, 1);
					}
                    if ((attachToIIS != null) && (toolsPopup != null))
                    {
                        attachToIIS.AddControl(toolsPopup.CommandBar, 2);
                    }
				}
				catch(System.ArgumentException)
				{
					//If we are here, then the exception is probably because a command with that name
					//  already exists. If so there is no need to recreate the command and we can 
                    //  safely ignore the exception.
				}
			}
		}

		/// <summary>Implements the OnDisconnection method of the IDTExtensibility2 interface. Receives notification that the Add-in is being unloaded.</summary>
		/// <param term='disconnectMode'>Describes how the Add-in is being unloaded.</param>
		/// <param term='custom'>Array of parameters that are host application specific.</param>
		/// <seealso class='IDTExtensibility2' />
		public void OnDisconnection(ext_DisconnectMode disconnectMode, ref Array custom)
		{
		}

		/// <summary>Implements the OnAddInsUpdate method of the IDTExtensibility2 interface. Receives notification when the collection of Add-ins has changed.</summary>
		/// <param term='custom'>Array of parameters that are host application specific.</param>
		/// <seealso class='IDTExtensibility2' />		
		public void OnAddInsUpdate(ref Array custom)
		{
		}

		/// <summary>Implements the OnStartupComplete method of the IDTExtensibility2 interface. Receives notification that the host application has completed loading.</summary>
		/// <param term='custom'>Array of parameters that are host application specific.</param>
		/// <seealso class='IDTExtensibility2' />
		public void OnStartupComplete(ref Array custom)
		{
		}

		/// <summary>Implements the OnBeginShutdown method of the IDTExtensibility2 interface. Receives notification that the host application is being unloaded.</summary>
		/// <param term='custom'>Array of parameters that are host application specific.</param>
		/// <seealso class='IDTExtensibility2' />
		public void OnBeginShutdown(ref Array custom)
		{
		}
		
		/// <summary>Implements the QueryStatus method of the IDTCommandTarget interface. This is called when the command's availability is updated</summary>
		/// <param term='commandName'>The name of the command to determine state for.</param>
		/// <param term='neededText'>Text that is needed for the command.</param>
		/// <param term='status'>The state of the command in the user interface.</param>
		/// <param term='commandText'>Text requested by the neededText parameter.</param>
		/// <seealso class='Exec' />
		public void QueryStatus(string commandName, vsCommandStatusTextWanted neededText, ref vsCommandStatus status, ref object commandText)
		{
			if(neededText == vsCommandStatusTextWanted.vsCommandStatusTextWantedNone)
			{
                if (commandName == _fullCommandName_SaveAndBuildExternal)
				{
					status = (vsCommandStatus)vsCommandStatus.vsCommandStatusSupported|vsCommandStatus.vsCommandStatusEnabled;
					return;
				}
                if (commandName == _fullCommandName_AttachToIIS)
				{
					status = (vsCommandStatus)vsCommandStatus.vsCommandStatusSupported|vsCommandStatus.vsCommandStatusEnabled;
					return;
				}
			}
		}

		/// <summary>Implements the Exec method of the IDTCommandTarget interface. This is called when the command is invoked.</summary>
		/// <param term='commandName'>The name of the command to execute.</param>
		/// <param term='executeOption'>Describes how the command should be run.</param>
		/// <param term='varIn'>Parameters passed from the caller to the command handler.</param>
		/// <param term='varOut'>Parameters passed from the command handler to the caller.</param>
		/// <param term='handled'>Informs the caller if the command was handled or not.</param>
		/// <seealso class='Exec' />
		public void Exec(string commandName, vsCommandExecOption executeOption, ref object varIn, ref object varOut, ref bool handled)
		{
			handled = false;
			if(executeOption == vsCommandExecOption.vsCommandExecOptionDoDefault)
			{
                if (commandName == _fullCommandName_SaveAndBuildExternal)
				{
                    handled = SaveAndBuildExternal();
				}
                if (commandName == _fullCommandName_AttachToIIS)
                {
                    handled = AttachToIIS();
                }
			}
		}

	    bool SaveAndBuildExternal()
	    {
	        _applicationObject.DTE.ExecuteCommand("File.SaveAll");

	        

	        _applicationObject.DTE.ExecuteCommand("Tools.ExternalCommand7");

	        return true;
	    }

        bool BuildCurrentSolution()
        {
            const string buildExecutable = @"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\MSBuild.exe";
            const string arguments = @"/m $(SolutionFileName) /v:n";
            const string runIn = _applicationObject.DTE.Solution.SolutionBuild.;

            var solutionFileName = _applicationObject.DTE.Solution.FileName;

            return false;
        }

	    bool AttachToIIS()
	    {
	        var processes = _applicationObject.Debugger.LocalProcesses;

	        var couldFindIIS = false;
	        foreach (Process process in processes)
	        {
	            if (process.Name.EndsWith("w3wp.exe"))
	            {
	                couldFindIIS = true;
	                process.Attach();
	            }
	        }

            if (!couldFindIIS)
            {
                MessageBox.Show("Could not find IIS, sorry!", "Not attached", MessageBoxButtons.OK);
            }
	        return couldFindIIS;
	    }

	    private DTE2 _applicationObject;
		private AddIn _addInInstance;
	}
}