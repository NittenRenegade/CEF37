using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Runtime.InteropServices;
using CefSharp;
using CefSharp.Wpf;
using ТестВК;

namespace WpfApp
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        dynamic Модуль1С;

        public class CallbackObjectForJs
        {
            public void showMessage(string msg)
            {//Read Note
                MessageBox.Show(msg);
            }
        }

        public MainWindow()
        {
            InitializeComponent();
            DataContext = new MainViewModel();

            browser.RegisterJsObject("callbackObj", new CallbackObjectForJs());
        }

        // For COM initialyze only
        public MainWindow(dynamic модуль1С, dynamic Object1C)
        {
            InitializeComponent();
            DataContext = new MainViewModel();

            Модуль1С = модуль1С;
            IExtWndsSupport n;
            n = (IExtWndsSupport)Object1C;
            IntPtr hwnd;
            n.GetAppMainFrame(out hwnd);

            var wih = new System.Windows.Interop.WindowInteropHelper(this);
            wih.Owner = hwnd;
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            Модуль1С.СообщитьСтр("Форма закрыта");
            Marshal.Release(Marshal.GetIDispatchForObject(Модуль1С));
            Marshal.ReleaseComObject(Модуль1С);
            Модуль1С = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            //browser.Address = "https://www.google.ru/maps";
        }
    }
}
