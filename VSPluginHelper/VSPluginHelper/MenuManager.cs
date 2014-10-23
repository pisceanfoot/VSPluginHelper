using System;
using System.Collections.Generic;
using System.Reflection;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.CommandBars;

namespace VSPluginHelper
{
    /// <summary>
    /// Add menu or command
    /// </summary>
    public class MenuManager
    {
        #region MEMBER VARIABLES

        ///////////////////////////////////////////////////////////////////////
        //
        // MEMBER VARIABLES
        //
        ///////////////////////////////////////////////////////////////////////
        //protected static readonly ILog _logger = LogManager.GetLogger((MethodBase.GetCurrentMethod().DeclaringType));
        private DTE2 application;

        private List<CommandBarEvents> menuItemHandlerList = new List<CommandBarEvents>();
        private Dictionary<string, CommandBase> cmdList = new Dictionary<string, CommandBase>();

        #endregion MEMBER VARIABLES

        #region INITIALIZATION

        ///////////////////////////////////////////////////////////////////////
        //
        // INITIALIZATION
        //
        ///////////////////////////////////////////////////////////////////////
        /// <summary>
        /// Initializes a new instance of the <see cref="T:MenuManager"/> class.
        /// </summary>
        /// <param name="application">The application.</param>
        public MenuManager(DTE2 application)
        {
            this.application = application;
        }

        #endregion INITIALIZATION

        #region MENU CREATION

        ///////////////////////////////////////////////////////////////////////
        //
        // MENU CREATION
        //
        ///////////////////////////////////////////////////////////////////////
        /// <summary>
        /// Creates the popup menu.
        /// </summary>
        /// <param name="commandBarName">Name of the command bar.</param>
        /// <param name="menuName">Name of the menu.</param>
        /// <param name="position"></param>
        /// <returns></returns>
        public CommandBarPopup CreatePopupMenu(string commandBarName, string menuName, int position)
        {
            CommandBar commandBar = GetCommandBar(commandBarName);
            if (commandBar != null)
            {
                try
                {
                    return commandBar.Controls[menuName] as CommandBarPopup;
                }
                catch
                {
                    CommandBarPopup menu = commandBar.Controls.Add(MsoControlType.msoControlPopup, Missing.Value, Missing.Value, position, true) as CommandBarPopup;
                    menu.Caption = menuName;
                    menu.TooltipText = "";

                    return menu;
                }
            }
            else
            {
                return null;
            }
        }

        public CommandBarPopup AddSubPopupMenu(CommandBarPopup commandBarPopup, string menuName, int position)
        {
            try
            {
                return commandBarPopup.CommandBar.Controls[menuName] as CommandBarPopup;
            }
            catch
            {
                CommandBarPopup menu = commandBarPopup.CommandBar.Controls.Add(MsoControlType.msoControlPopup, Missing.Value, Missing.Value, position, true) as CommandBarPopup;
                menu.Caption = menuName;
                menu.TooltipText = "";

                return menu;
            }
        }

        /// <summary>
        /// Creates the popup menu.
        /// </summary>
        /// <param name="popupMenu">The popup menu.</param>
        /// <param name="subPopupMenuName">Name of the sub popup menu.</param>
        /// <param name="position">The position.</param>
        public CommandBarPopup CreatePopupMenu(CommandBarPopup popupMenu, string subPopupMenuName, int position)
        {
            CommandBarPopup menu = (CommandBarPopup)popupMenu.Controls.Add(MsoControlType.msoControlPopup, Type.Missing, Type.Missing, position, true);
            menu.Caption = subPopupMenuName;
            menu.TooltipText = subPopupMenuName;
            return menu;
        }

        /// <summary>
        /// Creates the menu.
        /// </summary>
        /// <param name="commandBarName">Name of the command bar.</param>
        /// <param name="cmd">Name of the menu.</param>
        /// <returns></returns>
        public void CreateMenu(string commandBarName, CommandBase cmd)
        {
            CreateMenu(commandBarName, cmd, 1);
        }

        /// <summary>
        /// Creates the menu.
        /// </summary>
        /// <param name="commandBarName">Name of the command bar.</param>
        /// <param name="cmd"></param>
        /// <param name="position">ÃÌº”Œª÷√</param>
        /// <returns></returns>
        public void CreateMenu(string commandBarName, CommandBase cmd, int position)
        {
            if (GetCommandBar(commandBarName) != null)
            {
                CommandBarControl menuItem = GetCommandBar(commandBarName).Controls.Add(MsoControlType.msoControlButton, Missing.Value, Missing.Value, 1, true) as CommandBarControl;
                menuItem.Tag = cmd.Id.ToString();
                menuItem.Caption = cmd.Caption;
                menuItem.TooltipText = cmd.TooltipText;

                AddClickEventHandler(menuItem);
                AddCommandToList(cmd);
            }
        }

        /// <summary>
        /// Gets the command bar.
        /// </summary>
        /// <param name="commandBarName">Name of the command bar.</param>
        /// <returns></returns>
        private CommandBar GetCommandBar(string commandBarName)
        {
            try
            {
                return ((CommandBars)application.DTE.CommandBars)[commandBarName];
            }
            catch
            {
                return null;
            }
        }

        #endregion MENU CREATION

        #region MENU ITEMS CREATION

        ///////////////////////////////////////////////////////////////////////
        //
        // MENU ITEMS CREATION
        //
        ///////////////////////////////////////////////////////////////////////
        /// <summary>
        /// Adds the command.
        /// </summary>
        /// <param name="menu">The popup menu.</param>
        /// <param name="cmd">The command.</param>
        /// <param name="position">The position in the menu.</param>
        public void AddCommandMenu(CommandBarPopup popupMenu, CommandBase cmd, int position)
        {
            CommandBarControl menuItem = AddMenuItem(popupMenu, cmd, position);
            AddClickEventHandler(menuItem);
            AddCommandToList(cmd);
        }

        /// <summary>
        /// Add the menu item to the popup menu
        /// </summary>
        /// <param name="menu">The menu.</param>
        /// <param name="cmd">The CMD.</param>
        /// <param name="position">The position.</param>
        /// <returns></returns>
        private CommandBarControl AddMenuItem(CommandBarPopup menu, CommandBase cmd, int position)
        {
            cmd.Application = this.application;

            CommandBarControl menuItem = menu.Controls.Add(MsoControlType.msoControlButton, 1, "", position, true);
            menuItem.Tag = cmd.Id.ToString();
            menuItem.Caption = cmd.Caption;
            menuItem.TooltipText = cmd.TooltipText;
            return menuItem;
        }

        /// <summary>
        /// Adds handler to the menu item click event.
        /// </summary>
        /// <param name="menuItem">The menu item.</param>
        private void AddClickEventHandler(CommandBarControl menuItem)
        {
            CommandBarEvents menuItemHandler = (EnvDTE.CommandBarEvents)application.DTE.Events.get_CommandBarEvents(menuItem);
            menuItemHandler.Click += new _dispCommandBarControlEvents_ClickEventHandler(MenuItem_Click);
            menuItemHandlerList.Add(menuItemHandler);
        }

        /// <summary>
        /// Adds the command to list.
        /// </summary>
        /// <param name="cmd">The CMD.</param>
        private void AddCommandToList(CommandBase cmd)
        {
            if (!cmdList.ContainsKey(cmd.Id.ToString()))
            {
                cmdList.Add(cmd.Id.ToString(), cmd);
            }
        }

        #endregion MENU ITEMS CREATION

        #region MENU ITEMS HANDLER

        ///////////////////////////////////////////////////////////////////////
        //
        // MENU ITEMS HANDLER
        //
        ///////////////////////////////////////////////////////////////////////
        /// <summary>
        /// Handles the click on the menu item
        /// </summary>
        /// <param name="commandBarControl">The command bar control.</param>
        /// <param name="handled">if set to <c>true</c> [handled].</param>
        /// <param name="cancelDefault">if set to <c>true</c> [cancel default].</param>
        private void MenuItem_Click(object commandBarControl, ref bool handled, ref bool cancelDefault)
        {
            try
            {
                // We perform the command only if we found the command corresponding to the menu item clicked
                CommandBarControl menuItem = (CommandBarControl)commandBarControl;
                if (cmdList.ContainsKey(menuItem.Tag))
                {
                    CommandBase commandBase = cmdList[menuItem.Tag];
                    commandBase.Perform();
                }
            }
            catch (Exception e)
            {
                //_logger.ErrorFormat("An error occured : {0}", e.Message);
                //MessageBox.Show(e.Message, "MenuManager Explorer add in");
            }
        }

        #endregion MENU ITEMS HANDLER
    }
}