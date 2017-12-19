using System.Windows.Forms;
using System.Reflection;

namespace Favorites.TaskPane
{

    public partial class Settings : UserControl
    {

        public Settings()
        {
            InitializeComponent();
            this.pgdSettings.SelectedObject = Properties.Settings.Default;
        }

        public static void SetLabelColumnWidth(PropertyGrid grid, int width)
        {
            if (grid == null)
                return;

            FieldInfo fi = grid.GetType().GetField("gridView", BindingFlags.Instance | BindingFlags.NonPublic);
            if (fi == null)
                return;

            Control view = fi.GetValue(grid) as Control;
            if (view == null)
                return;

            MethodInfo mi = view.GetType().GetMethod("MoveSplitterTo", BindingFlags.Instance | BindingFlags.NonPublic);
            if (mi == null)
                return;
            mi.Invoke(view, new object[] { width });
        }

        private void pgdSettings_PropertyValueChanged(object s, PropertyValueChangedEventArgs e)
        {
            //Scripts.Ribbon.ribbonref.InvalidateRibbon();
        }

    }
}
