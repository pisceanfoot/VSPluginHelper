using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using EnvDTE;
using EnvDTE80;

namespace VSPluginHelper
{
    public class WindowManager
    {
        private DTE2 _applicationObject;
        private AddIn _addInInstance;

        public WindowManager(DTE2 _applicationObject, AddIn _addInInstance)
        {
            this._applicationObject =_applicationObject;
            this._addInInstance = _addInInstance;
        }

        public Window CreateAddinWindow<T>(string caption)
            where T : class
        {
            return this.CreateAddinWindow<T>(caption, Guid.NewGuid());
        }

        public Window CreateAddinWindow<T>(string caption, Guid id)
            where T : class
        {
            return CreateToolWindow<T>(caption, id, _addInInstance);
        }

        public Window CreateToolWindow<T>(string caption, Guid uiTypeGuid, AddIn addinInstance)
            where T : class
        {
            Type type = typeof(T);

            Windows2 win2 = _applicationObject.Windows as Windows2;
            if (win2 != null)
            {
                object controlObject = null;
                Window toolWindow = win2.CreateToolWindow2(addinInstance, type.Assembly.Location, type.FullName, caption, "{" + uiTypeGuid.ToString() + "}", ref controlObject);
                toolWindow.Visible = true;
                return toolWindow;
            }

            return null;
        }
    }
}
