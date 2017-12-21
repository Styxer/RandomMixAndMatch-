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
using Excel = Microsoft.Office.Interop.Excel;

namespace WpfApp1
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        private static Excel.Workbook MyBook = null;
        private static Excel.Application MyApp = null;
        private static Excel.Worksheet MySheet = null;
        private static int lastRow = 0;
        private static int lastColumn = 0;

        public static List<Person> personList = new List<Person>();
        static Random maleRand = new Random();
        static Random femaleRand = new Random();
        public static List<Coupls> couples = new List<Coupls>();

        public MainWindow()
        {
            InitializeComponent();


            string DB_PATH = @"C:\Users\ofir\Desktop\TestSheet.xlsx";
            try
            {
                InitializeExcel(DB_PATH);
                ReadMyExcel();
                mixAndMatch();
                WriteToExcel(couples);
                CloseExcel();

            }
            catch (Exception ex)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(MyApp);
                 throw;
            }
        }

        public static void CloseExcel()
        {
            MyBook.Saved = true;
            MyApp.Quit();

        }

        public static void WriteToExcel(List<Coupls> couples)
        {
            lastRow = 2 ;

            MySheet.Cells[lastRow, 3] = "maleName";
            MySheet.Cells[lastRow, 4] = "femaleName";
           
            for (int i = 0; i < couples.Count; i++)           
            {
                try
                {
                    lastRow ++;
                    MySheet.Cells[lastRow, 3] = couples[i].femaleName;
                    MySheet.Cells[lastRow, 4] = couples[i].maleName;
                 
                    // EmpList.Add(emp);
                    
                }
                catch (Exception ex)
                { } 
                finally
                {
                    MyBook.Save();
                }
            }

        }

        public void mixAndMatch()
        {
            List<Person> males = new List<Person>();
            List<Person> females = new List<Person>();
            foreach (var item in personList)
            {
                if (item.gender.ToLower() == "male")
                    males.Add(item);
                else
                    females.Add(item);
            }

            while (females.Any(x => x.gotMatch == false) && males.Any(x => x.gotMatch == false))
            {
                int choosenMale = maleRand.Next(0, males.Count);
                int choosenFemale = femaleRand.Next(0, females.Count);

                if (males[choosenMale].gotMatch == false && females[choosenFemale].gotMatch == false)
                {
                    males[choosenMale].gotMatch = true;
                    females[choosenFemale].gotMatch = true;
                    Coupls couple = new Coupls()
                    {
                        femaleName = females[choosenFemale].Name,
                        maleName = males[choosenMale].Name
                    };
                    if (!couples.Contains(couple))
                    {
                        couples.Add(couple);
                    }
                    females[choosenFemale].gotMatch = false;

                }  
            }
        }


        public static void InitializeExcel(string DB_PATH)
        {
            MyApp = new Excel.Application
            {
                Visible = false
            };
            MyBook = MyApp.Workbooks.Open(DB_PATH);
            MySheet = (Excel.Worksheet)MyBook.Sheets[1]; // Explict cast is not required here
            Excel.Range xlRange = MySheet.UsedRange;

          

            lastRow = xlRange.Rows.Count;
            lastColumn = xlRange.Columns.Count;
            //  lastRow = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;


        }


        public static List<Person> ReadMyExcel()
        {
            personList.Clear();
            for (int index = 2; index <= lastRow; index++)
            {
                System.Array MyValues = (System.Array)MySheet.get_Range("A" + index.ToString(), "B" + index.ToString()).Cells.Value;
                if (MyValues.GetValue(1, 1) != null && MyValues.GetValue(1, 2) != null)
                {
                    personList.Add(new Person
                    {

                        Name = MyValues.GetValue(1, 1).ToString(),
                        gender = MyValues.GetValue(1, 2).ToString(),
                        gotMatch = false
                    });

                }
                else
                    break;
            }
            return personList;




        }

        public class Person
        {
            public string Name { get; set; }
            public string age { get; set; }
            public string gender { get; set; }            
            public bool gotMatch { get; set; }
        }

        public class Coupls
        {
            public string maleName { get; set; }
            public string femaleName { get; set; }
         
        }

        private void Grid_Unloaded(object sender, RoutedEventArgs e)
        {

        }

        private void Window_Unloaded(object sender, RoutedEventArgs e)
        {

        }

        private void Window_Closed(object sender, EventArgs e)
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(MyApp);
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(MyApp);
        }
    }
}
