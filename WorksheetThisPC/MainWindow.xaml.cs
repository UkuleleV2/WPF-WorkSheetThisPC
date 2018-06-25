using System;
using System.IO;
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
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace WorksheetThisPC
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        StreamReader sr;
        string line;
        string source_dir;
        string destination_dir;
        List<Computer> list = new List<Computer>();
        public MainWindow()
        {
            InitializeComponent();
            source_dir = "";
            destination_dir = "";
        }

        public void overview_file()
        {
            string[] array2 = Directory.GetFiles(source_dir, "*.txt");
            foreach (string name in array2)
            {
                Console.WriteLine(name);
                try
                {   // Open the text file using a stream reader.
                    sr = new StreamReader(name);
                    searchfile();

                }
                catch (Exception e)
                {
                    Console.WriteLine("The file could not be read:" + name);
                    Console.WriteLine(e.Message);
                }
            }
        }
        public void searchfile()
        {
            string temp_string;
            bool ethernet_added = false;
            Computer temp = new Computer();
            while (!sr.EndOfStream)
            {
                line = sr.ReadLine();
                if (line.IndexOf("Host name:") != -1)
                {
                    temp_string = line.Substring(line.IndexOf(": ") + 2);
                    temp.hostname = temp_string;
                    Console.WriteLine(temp.hostname);
                }
                if (line.IndexOf("User name:") != -1)
                {
                    temp_string = line.Substring(line.IndexOf(": ") + 2);
                    temp.username = temp_string;
                    Console.WriteLine(temp.username);
                }
                if (line.IndexOf("Operating system:") != -1)
                {
                    temp_string = line.Substring(line.IndexOf(": ") + 2, line.IndexOf("(version") - line.IndexOf(": ") - 3);
                    temp.system = temp_string;
                    Console.WriteLine(temp_string);
                }
                if (line.IndexOf("Windows product key:") != -1)
                {
                    Regex regex = new Regex("([A-Z0-9]{5}-){4}[A-Z0-9]{5}$");
                    // Match the regular expression pattern against a text string.
                    Match match = regex.Match(line);
                    if (regex.IsMatch(line))
                    {

                        temp_string = line.Substring(match.Index);
                        temp.system_key = temp_string;
                        Console.WriteLine(temp_string);
                    }
                }

                if (line.IndexOf("Processor:") != -1)
                {
                    temp_string = line.Substring(line.IndexOf(": ") + 2, line.IndexOf("(architecture") - line.IndexOf(": ") - 3);
                    temp.processor = temp_string;
                    Console.WriteLine(temp_string);
                }
                if (line.IndexOf("Physical memory:") != -1)
                {
                    temp_string = line.Substring(line.IndexOf(": ") + 2);
                    temp.memory = temp_string;
                    Console.WriteLine(temp_string);
                }
                if (line.IndexOf("Disk:") != -1)
                {
                    temp_string = line.Substring(line.IndexOf(": ") + 2);
                    temp.disk = temp_string;
                    Console.WriteLine(temp.disk);
                }
                if (line.IndexOf("Network adapter:") != -1 && line.IndexOf("Ethernet") != -1 && ethernet_added != true)
                {
                    while (!sr.EndOfStream)
                    {
                        line = sr.ReadLine();
                        if (line.IndexOf("Adapter MAC-address:") != -1)
                        {
                            temp_string = line.Substring(line.IndexOf(": ") + 2);
                            temp.ethernet_mac = temp_string;
                            Console.WriteLine(temp.ethernet_mac);
                            ethernet_added = true;
                            break;
                        }

                    }
                }
                if (line.IndexOf("Network adapter:") != -1 && line.IndexOf("Wireless") != -1)
                {
                    while (!sr.EndOfStream)
                    {
                        line = sr.ReadLine();
                        if (line.IndexOf("Adapter MAC-address:") != -1)
                        {
                            temp_string = line.Substring(line.IndexOf(": ") + 2);
                            temp.wireless_mac = temp_string;
                            Console.WriteLine(temp.wireless_mac);
                            ethernet_added = true;
                            break;
                        }

                    }
                }
                if (line.IndexOf("Office") != -1)
                {
                    temp_string = line.Substring(0, line.IndexOf("-"));
                    temp.office = temp_string;
                    Console.WriteLine(temp.office);
                    Regex regex = new Regex("([A-Z0-9]{5}-){4}[A-Z0-9]{5}$");
                    // Match the regular expression pattern against a text string.
                    Match match = regex.Match(line);
                    if (regex.IsMatch(line))
                    {

                        // Console.WriteLine(line.Substring(match.Index,29));
                        temp.office_key = line.Substring(match.Index, 29);
                        Console.WriteLine(temp.office_key);
                    }
                }

            }
            list.Add(temp);
            sr.Close();
        }

        private void SelectFile(object sender, RoutedEventArgs e)
        {
            using (var dialog = new System.Windows.Forms.FolderBrowserDialog())
            {
                System.Windows.Forms.DialogResult result = dialog.ShowDialog();
                source_dir = dialog.SelectedPath;
            }
            selected_dir.Content = source_dir;
            overview_file();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (source_dir=="" && destination_dir == "")
            {
                System.Windows.MessageBox.Show("Source Folder or Destination file is not chosen!");
                return;
            }
            Microsoft.Office.Interop.Excel.Application oXL;
            Microsoft.Office.Interop.Excel._Workbook oWB;
            Microsoft.Office.Interop.Excel._Worksheet oSheet;
            Microsoft.Office.Interop.Excel.Range oRng;
            object misvalue = System.Reflection.Missing.Value;
            try
            {
                //Start Excel and get Application object.
                oXL = new Microsoft.Office.Interop.Excel.Application();
                oXL.Visible = true;

                //Get a new workbook.
                oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Add(""));
                oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;
                for(int i = 0;i<list.Capacity;i++)
                {
                    oSheet.Cells[i + 1, 1] = list[i].hostname;
                    oSheet.Cells[i + 1, 2] = list[i].username;
                    oSheet.Cells[i + 1, 3] = list[i].system;
                    oSheet.Cells[i + 1, 4] = list[i].system_key;
                    oSheet.Cells[i + 1, 5] = list[i].processor;
                    oSheet.Cells[i + 1, 6] = list[i].memory;
                    oSheet.Cells[i + 1, 7] = list[i].disk;
                    oSheet.Cells[i + 1, 8] = list[i].ethernet_mac;
                    oSheet.Cells[i + 1, 9] = list[i].wireless_mac;
                    oSheet.Cells[i + 1, 10] = list[i].office;
                    oSheet.Cells[i + 1, 11] = list[i].office_key;
                }

                //Add table headers going cell by cell.

                oXL.Visible = false;
                oXL.UserControl = false;
                oWB.SaveAs(destination_dir, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                    false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                oWB.Close();

            }
            catch (Exception ex)
            {
                Console.WriteLine("Cannot write to excel file");
                Console.WriteLine(ex.Message);
            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            // Configure save file dialog box
            Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();
            dlg.FileName = "Worksheet file"; // Default file name
            dlg.DefaultExt = ".xlx"; // Default file extension
            dlg.Filter = "Excel files (.xlx)|*.xlx"; // Filter files by extension

            // Show save file dialog box
            Nullable<bool> result = dlg.ShowDialog();

            // Process save file dialog box results
            if (result == true)
            {
                // Save document
                destination_dir = dlg.FileName;
                excel_file.Content = destination_dir;
            }
        }
    }
}
