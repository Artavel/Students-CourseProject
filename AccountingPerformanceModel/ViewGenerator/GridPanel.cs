using System;
using System.Windows.Forms;

namespace ViewGenerator
{
    // Пользовательская панель
    public class GridPanel: UserControl
    {
        public event Action<object> GridSelectedChanged = delegate { };

        public void OnGridSelectedChanged(object userClass)
        {
            GridSelectedChanged(userClass);
        }
    }
}
