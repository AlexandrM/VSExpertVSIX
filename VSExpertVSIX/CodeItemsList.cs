﻿//------------------------------------------------------------------------------
// <copyright file="CodeItemsList.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.ComponentModel.Design;
using System.Globalization;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using ASEExpertVS2005.CodeItemsList;
using EnvDTE;
using EnvDTE80;

namespace VSExpertVSIX
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class CodeItemsList
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0301;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("eee331c3-05bb-4316-a339-85c107e23240");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;
        public static DTE2 DTE;
        /// <summary>
        /// Initializes a new instance of the <see cref="CodeItemsList"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private CodeItemsList(Package package)
        {
            if (package == null)
            {
                throw new ArgumentNullException("package");
            }

            this.package = package;

            DTE = Package.GetGlobalService(typeof(SDTE)) as DTE2;

            OleMenuCommandService commandService = this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                var menuCommandID = new CommandID(CommandSet, CommandId);
                var menuItem = new MenuCommand(this.MenuItemCallback, menuCommandID);
                commandService.AddCommand(menuItem);
            }
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static CodeItemsList Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static void Initialize(Package package)
        {
            Instance = new CodeItemsList(package);
        }

        private FmCodeItems fmCodeItems = new FmCodeItems();
        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void MenuItemCallback(object sender, EventArgs e)
        {
            var filter = "";
            if ((DTE != null) && (DTE.ActiveDocument != null) && (DTE.ActiveDocument.Selection != null))
            {
                var textSelection = (DTE.ActiveDocument.Selection as TextSelection);
                if (textSelection != null)
                {
                    filter = textSelection.Text;
                }
            }
            fmCodeItems.DoDialog(filter);
        }
    }
}
