using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Tekla.Structures;
using Tekla.Structures.Model;
using Tekla.Structures.Geometry3d;

namespace vtk_final
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog fdlg = new OpenFileDialog();
            fdlg.Title = "C# Corner Open File Dialog";
            fdlg.InitialDirectory = @"c:\";
            fdlg.Filter = "All files (*.*)|*.*|All files (*.*)|*.*";
            fdlg.FilterIndex = 2;
            fdlg.RestoreDirectory = true;
            if (fdlg.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = fdlg.FileName;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ArrayList No_list = new ArrayList();
            ArrayList Ele_list = new ArrayList();
            ArrayList Hsppm_list = new ArrayList();

            string text = System.IO.File.ReadAllText(textBox1.Text);

            string[] lines = System.IO.File.ReadAllLines(textBox1.Text);

            foreach (string line in lines)
            {
                if(line.StartsWith("NO"))
                {
                    No_list.Add(line);
                }
                else if(line.StartsWith("ELE"))
                {
                    Ele_list.Add(line);
                }
                else if (line.StartsWith("HSPPM"))
                {
                    Hsppm_list.Add(line);
                }
            }

            List<Tekla.Structures.Geometry3d.Point> mypointlist = new List<Tekla.Structures.Geometry3d.Point>();

            Model mymodel = new Model();
            if(mymodel.GetConnectionStatus())
            {
                ControlPoint mycontrol_point = new ControlPoint();
                ArrayList mylist2 = new ArrayList();
                foreach (string item in No_list)
                {
                    var stringarray = item.Split(',');
                    Tekla.Structures.Geometry3d.Point samplepoint = new Tekla.Structures.Geometry3d.Point();
                    samplepoint.X = Convert.ToDouble(stringarray[1])*1000;
                    samplepoint.Y = Convert.ToDouble(stringarray[2])*1000;
                    samplepoint.Z = Convert.ToDouble(stringarray[3])*1000;
                    mypointlist.Add(samplepoint);
                }

                int q = 0;

                foreach (string item in Ele_list)
                {
                    var stringarray = item.Split(',');

                    string first_input_point = new String(stringarray[1].Where(Char.IsDigit).ToArray());
                    string second_input_point = new String(stringarray[2].Where(Char.IsDigit).ToArray());

                    Beam mybeam = new Beam();
                    mybeam.StartPoint = mypointlist[Convert.ToInt32(first_input_point) - 1];
                    mybeam.EndPoint = mypointlist[Convert.ToInt32(second_input_point) - 1];
                    mybeam.Name = stringarray[0];
                    mybeam.Class = (q+1).ToString();
                    string profile = "Ø" + stringarray[3].Replace(" ","") + "*" + stringarray[4].Replace(" ", "");
                    mybeam.Profile.ProfileString = profile.ToString();
                    mybeam.Position.Plane = Position.PlaneEnum.MIDDLE;
                    mybeam.Position.Rotation = Position.RotationEnum.TOP;
                    mybeam.Position.Depth = Position.DepthEnum.MIDDLE;
                    mybeam.Insert();

                    q = q + 1;
                }

                int a = 1;

                foreach (string item in Hsppm_list)
                {
                    var stringarray = item.Split(',');

                    Tekla.Structures.Geometry3d.Point samplepoint = new Tekla.Structures.Geometry3d.Point();
                    samplepoint.X = Convert.ToDouble(stringarray[3]) * 1000;
                    samplepoint.Y = Convert.ToDouble(stringarray[4]) * 1000;
                    samplepoint.Z = Convert.ToDouble(stringarray[5]) * 1000;

                    Beam mybeam1 = new Beam();

                    double distance = Convert.ToDouble(stringarray[6]) * 100 / 2;

                    mybeam1.StartPoint.X = Convert.ToDouble(samplepoint.X) - distance;
                    mybeam1.StartPoint.Y = Convert.ToDouble(samplepoint.Y);
                    mybeam1.StartPoint.Z = Convert.ToDouble(samplepoint.Z);

                    mybeam1.EndPoint.X = Convert.ToDouble(samplepoint.X) + distance;
                    mybeam1.EndPoint.Y = Convert.ToDouble(samplepoint.Y);
                    mybeam1.EndPoint.Z = Convert.ToDouble(samplepoint.Z);


                    mybeam1.Name = stringarray[0] + stringarray[1] + stringarray[6];
                    mybeam1.Class = a.ToString();
                    string profile = "SPHERE"+Convert.ToDouble(stringarray[6])*100 ;
                    mybeam1.Profile.ProfileString = profile.ToString();
                    mybeam1.Position.Plane = Position.PlaneEnum.MIDDLE;
                    mybeam1.Position.Rotation = Position.RotationEnum.TOP;
                    mybeam1.Position.Depth = Position.DepthEnum.MIDDLE;
                    mybeam1.Insert();

                    a = a + 1;
                }



                //MessageBox.Show("asd");
            }
            mymodel.CommitChanges();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Model mymodel = new Model();
            var selector = new Tekla.Structures.Model.UI.ModelObjectSelector();
            var myenum = selector.GetSelectedObjects();

            foreach (Beam item in myenum)
            {
                string name1 = item.Name;
                string startpoint1 = (item.StartPoint).ToString();
                string endpoint1 = (item.EndPoint).ToString();

                var split = name1.Split(' ');

                string abhi = name1 + System.Environment.NewLine + startpoint1 + System.Environment.NewLine + endpoint1;

                //MessageBox.Show(abhi);

                Tekla.Structures.Model.UI.GraphicsDrawer mygraphic = new Tekla.Structures.Model.UI.GraphicsDrawer();
                mygraphic.DrawText(item.StartPoint,abhi,new Tekla.Structures.Model.UI.Color(0.35,0.35,0.35));

            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Model mymodel = new Model();
            var selector = new Tekla.Structures.Model.UI.ModelObjectSelector();
            var myenum = selector.GetSelectedObjects();

            ArrayList mylist = new ArrayList();

            foreach (Beam item in myenum)
            {
                string name1 = item.Name;
                string startpoint1 = (item.StartPoint).ToString();
                string endpoint1 = (item.EndPoint).ToString();

                var split = name1.Split(' ');

                string abhi = name1 + System.Environment.NewLine + startpoint1 + System.Environment.NewLine + endpoint1;

                mylist.Add(abhi);
            }


            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            if (excelApp != null)
            {
                Microsoft.Office.Interop.Excel.Workbook excelWorkbook = excelApp.Workbooks.Add();
                Microsoft.Office.Interop.Excel.Worksheet excelWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)excelWorkbook.Sheets.Add();

                int increase = 1;

                foreach (string item in mylist)
                {
                    excelWorksheet.Cells[increase, 1] = item;
                    increase = increase + 1;
                }


                excelApp.ActiveWorkbook.SaveAs(@"C:\Users\absa\Downloads\geometry\testrun.xls", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal);

                excelWorkbook.Close();
                excelApp.Quit();

                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelWorksheet);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelWorkbook);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelApp);
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
    }
}
