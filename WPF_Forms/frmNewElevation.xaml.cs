﻿using System;
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
using System.Windows.Shapes;

namespace LifestyleDesign_r24
{
    /// <summary>
    /// Interaction logic for frmNewElevation.xaml
    /// </summary>
    public partial class frmNewElevation : Window
    {
        public frmNewElevation()
        {
            InitializeComponent();

            List<string> listElevations = new List<string> { "A", "B", "C", "D", "S", "T" };

            foreach (string elevation in listElevations)
            {
                cmbElevation.Items.Add(elevation);
            }

            cmbElevation.SelectedIndex = 0;

            List<string> listFloors = new List<string> { "1", "2", "3" };

            foreach (string floor in listFloors)
            {
                cmbFloors.Items.Add(floor);
            }

            cmbFloors.SelectedIndex = 0;

            List<string> listFoundations = new List<string> { "Basement", "Crawlspace", "Slab" };

            foreach (string foundation in listFoundations)
            {
                cmbFoundation.Items.Add(foundation);
            }

            cmbFoundation.SelectedIndex = 0;
        }

        internal string GetComboboxElevation()
        {
            return cmbElevation.Text.ToString();
        }

        internal string GetComboboxFloors()
        {
            return cmbFloors.SelectedItem.ToString();
        }

        internal string GetComboboxFoundation()
        {
            return cmbFoundation.SelectedItem.ToString();
        }

        public string Worksheet()
        {
            return cmbFoundation.SelectedItem.ToString() + '-' + cmbFloors.SelectedItem.ToString();
        }

        private void btnOK_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = true;
            this.Close();
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }

        private void btnHelp_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
