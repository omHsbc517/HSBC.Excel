using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace HSBC.InsuranceDataAnalysis.UI
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : SFE.Theme.Window
    {
        public MainWindow()
        {
            InitializeComponent();

            this.ApplicationVersion = "0.0.0.1";
            //((MainWindowViewModel)DataContext).CurrentUser = this.UserName;

            //Closing += (r, s) => { ((MainWindowViewModel)DataContext).PrepareStopProcess(); };
        }

        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);

            GlassHelper.ExtendGlassFrame(this, new Thickness(-1));
        }
    }

}
