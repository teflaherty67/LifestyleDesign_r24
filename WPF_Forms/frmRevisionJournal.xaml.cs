﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
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
using System.Windows.Shapes;

namespace LifestyleDesign_r24
{
    /// <summary>
    /// Interaction logic for frmRevisionJournal.xaml
    /// </summary>
    public partial class frmRevisionJournal : Window
    {
        string folderPath = "";
        string journalFilePath = "";
        string searchPath1 = @"S:\Shared Folders\-Job Log-\01-Current Jobs";
        string searchPath2 = @"S:\Shared Folders\-Job log-\02-Completed Jobs";

        BindingList<clsJournalData> journalDataList = new BindingList<clsJournalData>();
        clsJournalData curEdit;

        public frmRevisionJournal(string filePath)
        {
            InitializeComponent();

            string[] directory1 = Directory.GetDirectories(searchPath1);
            string[] directory2 = Directory.GetDirectories(searchPath2);

            // search path 1

            foreach (string dir in directory1)
            {
                if (dir.Contains(GlobalVars.JobNumber))
                    folderPath = dir;
            }

            // search path 2

            if (folderPath == "")
            {
                foreach (string dir in directory2)
                {
                    if (dir.Contains(GlobalVars.JobNumber))
                        folderPath = dir;
                }
            }

            if (folderPath == "")
            {
                folderPath = searchPath1;
            }

            tbFileNmae.Text = System.IO.Path.GetFileName(filePath);

            string curFileName = GlobalVars.JobNumber + "_Log.txt";

            journalFilePath = folderPath + @"\" + curFileName;

            ReadToDoFile();
        }

        private void ReadToDoFile()
        {
            if (File.Exists(journalFilePath))
            {
                int counter = 0;
                string[] strings = File.ReadAllLines(journalFilePath);

                foreach (string line in strings)
                {
                    clsJournalData curToDo = new clsJournalData(counter + 1, line);
                    journalDataList.Add(curToDo);
                    counter++;
                }
            }

            ShowData();
        }

        private void ShowData()
        {
            lbxTasks.ItemsSource = null;
            lbxTasks.ItemsSource = journalDataList;
            lbxTasks.DisplayMemberPath = "Display";
        }

        private void btnAddEdit_Click(object sender, RoutedEventArgs e)
        {

            if (curEdit == null)
            {
                AddToDoItem(tbxItem.Text);
            }
            else
            {
                CompleteEditingItem();
            }

            tbxItem.Text = "";
        }

        private void AddToDoItem(string todoText)
        {
            clsJournalData curToDo = new clsJournalData(journalDataList.Count + 1, todoText);
            journalDataList.Add(curToDo);

            WriteToDoFile();
        }

        private void WriteToDoFile()
        {
            using (StreamWriter writer = File.CreateText(journalFilePath))
            {
                foreach (clsJournalData curToDo in lbxTasks.Items)
                {
                    curToDo.UpdateDisplayString();
                    writer.WriteLine(curToDo.Display);
                }
            }

            ShowData();
        }

        private void CompleteEditingItem()
        {
            foreach (clsJournalData todo in journalDataList)
            {
                if (todo == curEdit)
                    todo.Text = tbxItem.Text;
            }

            curEdit = null;
            tbkAddEdit.Text = "Add Item";
            btnAddEdit.Content = "Add Item";

            WriteToDoFile();
        }

        private void btnEdit_Click(object sender, RoutedEventArgs e)
        {
            if (lbxTasks.SelectedItems != null)
            {
                clsJournalData curToDo = lbxTasks.SelectedItem as clsJournalData;
                StartEditingItem(curToDo);
            }
        }

        private void StartEditingItem(clsJournalData curToDo)
        {
            curEdit = curToDo;

            tbkAddEdit.Text = "Update Item";
            btnAddEdit.Content = "Update Item";
            tbxItem.Text = curToDo.Text;
        }

        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            if (lbxTasks.SelectedItems != null)
            {
                clsJournalData curToDo = lbxTasks.SelectedItem as clsJournalData;
                RemoveItem(curToDo);
            }
        }

        private void RemoveItem(clsJournalData curToDo)
        {
            journalDataList.Remove(curToDo);
            ReOrderToDoItems();

            WriteToDoFile();
        }

        private void ReOrderToDoItems()
        {
            for (int i = 0; i < journalDataList.Count; i++)
            {
                journalDataList[i].PositionNumber = i + 1;
                journalDataList[i].UpdateDisplayString();
            }
            WriteToDoFile();
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
