using RPA_SummerProj.core.module;
using System;
using System.Activities;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RPA_SummerProj
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Example_Click(object sender, EventArgs e)
        {
            #region synchronize Activities
            AutoResetEvent createEvent = new AutoResetEvent(false);
            AutoResetEvent openEvent = new AutoResetEvent(false);
            AutoResetEvent appendEvent = new AutoResetEvent(false);
            AutoResetEvent saveEvent = new AutoResetEvent(false);
            AutoResetEvent writeEvent = new AutoResetEvent(false);
            AutoResetEvent writeEvent2 = new AutoResetEvent(false);
            AutoResetEvent readEvent = new AutoResetEvent(false);
            AutoResetEvent readEvent2 = new AutoResetEvent(false);
            AutoResetEvent closeEvent = new AutoResetEvent(false);
            AutoResetEvent syncEvent = new AutoResetEvent(false);
            #endregion

            //Console.WriteLine("Please Enter the instance Name ");
            //string instanceName = Console.ReadLine();
            string instanceName = "Excel";
            Program instance = new Program();

            #region Sample data and inputs
            object[,] data = { { 1, 2, 3 } };
            var openInput = new Dictionary<string, object>() { { "Path", @"C:\test.xlsx" },
                                                            { "SheetName", null}, { "instance", this}, { "instanceName", instanceName} };
            var commonInput = new Dictionary<string, object>() { { "instance", instance }, { "instanceName", instanceName } };
            var saveInput = new Dictionary<string, object>() { { "instance", instance }, { "instanceName", instanceName }, { "Path", null } };
            var writeInput = new Dictionary<string, object>() { { "instance", instance}, { "instanceName", instanceName}, { "SheetName", "Sheet1"},
                                                                { "Cell", "A1"}, { "Data", "Hello"} };
            var writeInput2 = new Dictionary<string, object>() { { "instance", instance}, { "instanceName", instanceName}, { "SheetName", "Sheet1"},
                                                                { "Range", "A2:C2"}, { "Data", data } };
            var readInput = new Dictionary<string, object>() { { "instance", instance}, { "instanceName", instanceName}, { "SheetName", "Sheet1"},
                                                                { "Cell", "A1"} };
            var readInput2 = new Dictionary<string, object>() { { "instance", instance}, { "instanceName", instanceName}, { "SheetName", "Sheet1"},
                                                                { "Range", "A2:C2"} };

            #endregion
            /*
            WorkflowApplication wfApp = new WorkflowApplication(new Activity1(), input);
            wfApp.Completed = delegate (WorkflowApplicationCompletedEventArgs e)
            {
                Console.WriteLine("Completed");
                Debug.WriteLine("What");
                syncEvent.Set();
            };
            wfApp.Run();
            syncEvent.WaitOne();
            */

            #region Excel Activities Example
            WorkflowApplication create = new WorkflowApplication(new ExcelCreate(), commonInput);
            WorkflowApplication open = new WorkflowApplication(new ExcelOpen(), openInput);
            WorkflowApplication append = new WorkflowApplication(new ExcelAppend(), commonInput);
            WorkflowApplication write_cell = new WorkflowApplication(new ExcelWrite_Cell(), writeInput);
            WorkflowApplication write_range = new WorkflowApplication(new ExcelWrite_Range(), writeInput2);
            WorkflowApplication read_cell = new WorkflowApplication(new ExcelRead_Cell(), readInput);
            WorkflowApplication read_range = new WorkflowApplication(new ExcelRead_Range(), readInput2);
            WorkflowApplication save = new WorkflowApplication(new ExcelSave(), saveInput);
            WorkflowApplication close = new WorkflowApplication(new ExcelClose(), commonInput);

            create.Completed = delegate (WorkflowApplicationCompletedEventArgs ce)
            {
                Console.WriteLine("Created Completed");
                createEvent.Set();
            };

            open.Completed = delegate (WorkflowApplicationCompletedEventArgs oe)
            {
                Console.WriteLine("Open Completed");
                openEvent.Set();
            };

            append.Completed = delegate (WorkflowApplicationCompletedEventArgs ae)
            {
                Console.WriteLine("Append Completed");
                appendEvent.Set();
            };

            write_cell.Completed = delegate (WorkflowApplicationCompletedEventArgs we)
            {
                Console.WriteLine("Write Completed");
                writeEvent.Set();
            };

            write_range.Completed = delegate (WorkflowApplicationCompletedEventArgs we2)
            {
                Console.WriteLine("Write2 Completed");
                writeEvent2.Set();
            };

            read_cell.Completed = delegate (WorkflowApplicationCompletedEventArgs re)
            {
                Console.WriteLine("Read Completed");
                object get_data = re.Outputs["Data"];
                Console.WriteLine(get_data);
                readEvent.Set();
            };

            read_range.Completed = delegate (WorkflowApplicationCompletedEventArgs re2)
            {
                Console.WriteLine("Read2 Completed");
                object[,] sample = (object[,])re2.Outputs["Data"];
                foreach (var item in sample)
                {
                    Console.WriteLine(item);
                }
                readEvent2.Set();
            };

            save.Completed = delegate (WorkflowApplicationCompletedEventArgs se)
            {
                Console.WriteLine("Save Completed");
                saveEvent.Set();
            };



            close.Completed = delegate (WorkflowApplicationCompletedEventArgs ce)
            {
                Console.WriteLine("Close Completed");
                closeEvent.Set();
            };

            create.Run();
            createEvent.WaitOne();
            //open.Run();
            //openEvent.WaitOne();
            append.Run();
            appendEvent.WaitOne();
            write_cell.Run();
            writeEvent.WaitOne();
            write_range.Run();
            writeEvent2.WaitOne();
            read_cell.Run();
            readEvent.WaitOne();
            read_range.Run();
            readEvent2.WaitOne();
            save.Run();
            saveEvent.WaitOne();
            close.Run();
            closeEvent.WaitOne();
            #endregion
        }
    }
}
