using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.IO;
using System.Threading;
using System.Collections.Specialized;
using System.Collections;
using System.Windows.Interop;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.ProcessPower.PnIDObjects;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace ListPipelineObjects
{
    /// <summary>
    /// Interaction logic for ListPipelineObjectsUI.xaml
    /// </summary>
    public partial class ListPipelineObjectsUI : Window
    {
        public ListPipelineObjectsUI()
        {
            InitializeComponent();
        }

        private void Window_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            DragMove();
        }

        public void CleanString()
        {
            ListLinesObjsResult.Text = String.Empty;
        }

        public void AppendString(String msg)
        {
            ListLinesObjsResult.Text = ListLinesObjsResult.Text + msg;
        }

        public void AppendNewLineString(String msg)
        {
            ListLinesObjsResult.Text = ListLinesObjsResult.Text + Environment.NewLine + msg;
        }

        private void CloseBtn_Clicked(object sender, RoutedEventArgs e)
        {
            HideWindow();
        }

        private void HideWindow()
        {
            this.Visibility = System.Windows.Visibility.Hidden;
        }

        private void Window_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                HideWindow();
            }
        }
    }

    class ListPipelineObjectsUIHelpler
    {
        private static ListPipelineObjectsUI listPipelineObjs = null;
        public static void ShowWindow(String msg)
        {
            if (listPipelineObjs == null)
            {
                listPipelineObjs = new ListPipelineObjectsUI();
            }

            WindowInteropHelper wih = new WindowInteropHelper(listPipelineObjs);
            wih.Owner = AcadApp.MainWindow.Handle;

            listPipelineObjs.CleanString();
            listPipelineObjs.AppendString(msg);
            listPipelineObjs.Visibility = System.Windows.Visibility.Visible;
            listPipelineObjs.ShowInTaskbar = false;
            listPipelineObjs.Show();
        }
    }
}
