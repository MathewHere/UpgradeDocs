using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using NLog.Fluent;
using WebSupergoo.WordGlue;

namespace UpgradeDocs
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
       
    public MainWindow()
        {
            InitializeComponent();
        }

        private  void Button_Click(object sender, RoutedEventArgs e)
        {
           
            BusinessClass bo=new BusinessClass();
            //For the display of operation progress to UI.
            lbComplete.Visibility = Visibility.Visible;
            btnProcess.Content = "Processing!";
            btnProcess.IsEnabled = false;
            List<string > filedToprocess=new List<string>();
            bo.DirSearch( DirPath.Text, filedToprocess);
            lblTotalFiles.Content = filedToprocess.Count;
            bo.UpgradeOfficeFiles(filedToprocess, Word,Excel,Ppt,lbComplete,btnProcess);
        }
        // ************************


    }
}
