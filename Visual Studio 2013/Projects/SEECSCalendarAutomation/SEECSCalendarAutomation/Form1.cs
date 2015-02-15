using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SEECSCalendarAutomation
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

            if (comboBox1.Text == "null" || comboBox1.Text == "" || comboBox2.Text == "null" || comboBox2.Text == "" || textBox1.Text == "null" || textBox1.Text == "" || textBox2.Text == "null" || textBox2.Text == "")
            {
                MessageBox.Show("Please enter ALL information before saving");
                return;
            }

            string start_time = "null";
            string end_time = "null";

            string path = @"C:\Users\MMA\Desktop\";

            string date = "null";
            

            switch (comboBox1.SelectedItem.ToString())
            {
                case "1":
                    start_time = "9:00 AM";
                    end_time = "9:50 AM";
                    break;

                case "2":
                    start_time = "10:00 AM";
                    end_time = "10:50 AM";
                    break;

                case "3":
                    start_time = "11:00 AM";
                    end_time = "11:50 AM";
                    break;

                case "4":
                    start_time = "12:00 PM";
                    end_time = "12:50 PM";
                    break;

                case "5":
                    start_time = "2:00 PM";
                    end_time = "2:50 PM";
                    break;

                case "6":
                    start_time = "3:00 PM";
                    end_time = "3:50 PM";
                    break;

                case "7":
                    start_time = "4:00 AM";
                    end_time = "4:50 PM";
                    break;

            }

            switch (comboBox2.SelectedItem.ToString())
            {
                case "Monday":

                    DateTime today = DateTime.Today;
                    // The (... + 7) % 7 ensures we end up with a value in the range [0, 6]
                    int daysUntilMonday = ((int)DayOfWeek.Monday - (int)today.DayOfWeek + 7) % 7;
                    DateTime nextMonday = today.AddDays(daysUntilMonday);

                    date = nextMonday.ToShortDateString();

                    break;

                case "Tuesday":
                    today = DateTime.Today;
                    // The (... + 7) % 7 ensures we end up with a value in the range [0, 6]
                    int daysUntilTuesday = ((int)DayOfWeek.Tuesday - (int)today.DayOfWeek + 7) % 7;
                    DateTime nextTuesday = today.AddDays(daysUntilTuesday);

                    date = nextTuesday.ToShortDateString();
                    break;

                case "Wednesday":
                    today = DateTime.Today;
                    // The (... + 7) % 7 ensures we end up with a value in the range [0, 6]
                    int daysUntilWednesday = ((int)DayOfWeek.Wednesday - (int)today.DayOfWeek + 7) % 7;
                    DateTime nextWednesday = today.AddDays(daysUntilWednesday);

                    date = nextWednesday.ToShortDateString();
                    break;

                case "Thursday":
                    today = DateTime.Today;
                    // The (... + 7) % 7 ensures we end up with a value in the range [0, 6]
                    int daysUntilThursday = ((int)DayOfWeek.Thursday - (int)today.DayOfWeek + 7) % 7;
                    DateTime nextThursday = today.AddDays(daysUntilThursday);

                    date = nextThursday.ToShortDateString();
                    break;
                    
                case "Friday":
                    today = DateTime.Today;
                    // The (... + 7) % 7 ensures we end up with a value in the range [0, 6]
                    int daysUntilFriday = ((int)DayOfWeek.Friday - (int)today.DayOfWeek + 7) % 7;
                    DateTime nextFriday = today.AddDays(daysUntilFriday);

                    date = nextFriday.ToShortDateString();
                    break;


            }

            //Subject,Start Date,Start Time,End Date,End Time,All Day Event
            string text = textBox1.Text + ',' + date + ',' + start_time + ',' + date + ',' + end_time + ',' + "False";

            // create/append to file if file already exist

            if (File.Exists(path + textBox2.Text + ".csv"))
            {
                using (System.IO.StreamWriter file = new System.IO.StreamWriter(path + textBox2.Text + ".csv", true))
                {
                    file.WriteLine(text);
                }
            }
            else
            {
                using (System.IO.StreamWriter file = new System.IO.StreamWriter(path + textBox2.Text + ".csv", true))
                {
                    file.WriteLine("Subject,Start Date,Start Time,End Date,End Time,All Day Event");
                    file.WriteLine(text);
                }
            }

            comboBox1.Text = "";
            comboBox2.Text = "";
            textBox1.Text = "";
            textBox2.Text = "";

            MessageBox.Show("File Saved Successfully");


        }

        private void button2_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
    }
}
