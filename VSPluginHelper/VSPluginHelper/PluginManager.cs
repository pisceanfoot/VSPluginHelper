using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EnvDTE;
using EnvDTE80;

namespace VSPluginHelper
{
    public class PluginManager
    {
        private static PluginManager pluginManager;

        private DTE2 _applicationObject;
        private AddIn _addInInstance;

        private MenuManager menuManager;
        private WindowManager windowManager;

        private PluginManager(DTE2 _applicationObject, AddIn _addInInstance)
        {
            this._applicationObject =_applicationObject;
            this._addInInstance = _addInInstance;

            menuManager = new MenuManager(_applicationObject);
            windowManager = new WindowManager(_applicationObject, _addInInstance);
        }

        public static void Init(DTE2 _applicationObject, AddIn _addInInstance)
        {
            pluginManager = new PluginManager(_applicationObject, _addInInstance);
        }

        public static PluginManager Current
        {
            get
            {
                return pluginManager;
            }
        }

        public MenuManager MenuManager 
        {
            get
            {
                return this.menuManager;
            } 
        }

        public WindowManager WindowManager
        {
            get
            {
                return this.windowManager;
            }
        }
    }
}
