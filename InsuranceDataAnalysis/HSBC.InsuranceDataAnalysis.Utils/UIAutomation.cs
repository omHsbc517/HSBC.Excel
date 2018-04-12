////===============================================================================
//
//  Copyright © 2018 中软国际.HSBC业务线.第二事业部.保险与卡交付部 All rights reserved    
//  
//  Filename :UIAutomation
//  Description :
//
//  Created by Tina at  2/24/2018 1:21:58 PM
//
////===============================================================================
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Automation;

namespace HSBC.InsuranceDataAnalysis.Utils
{
    public static class UIAutomation
    {
        private delegate bool WNDENUMPROC(IntPtr hWnd, int lParam);

        // enumeration all windows
        [DllImport("user32.dll")]
        private static extern bool EnumWindows(WNDENUMPROC lpEnumFunc, int lParam);

        // get window text 
        [DllImport("user32.dll")]
        private static extern int GetWindowTextW(IntPtr hWnd, [MarshalAs(UnmanagedType.LPWStr)]StringBuilder lpString, int nMaxCount);

        // get window class name
        [DllImport("user32.dll")]
        private static extern int GetClassNameW(IntPtr hWnd, [MarshalAs(UnmanagedType.LPWStr)]StringBuilder lpString, int nMaxCount);

        // self difine class
        private struct WindowInfo
        {
            public IntPtr hWnd;
            public string szWindowName;
            public string szClassName;
        }

        private static IntPtr GetHandle(string title, string className)
        {
            IntPtr pt = new IntPtr();

            //enum all desktop windows 
            EnumWindows(delegate (IntPtr hWnd, int lParam)
            {
                WindowInfo wnd = new WindowInfo();
                StringBuilder sb = new StringBuilder(256);

                //get hwnd 
                wnd.hWnd = hWnd;

                //get window name  
                GetWindowTextW(hWnd, sb, sb.Capacity);
                wnd.szWindowName = sb.ToString();

                //get window class 
                GetClassNameW(hWnd, sb, sb.Capacity);
                wnd.szClassName = sb.ToString();

                if (wnd.szWindowName.Equals(title) && wnd.szClassName.Equals(className))
                {
                    pt = hWnd;
                    return false;
                }

                return true;
            }, 0);

            return pt;
        }

        public static AutomationElement GetWnd(string title, string className)
        {
            IntPtr pt = GetHandle(title, className);
            if (pt == IntPtr.Zero || pt == null)
            {
                return null;
            }
            return AutomationElement.FromHandle(pt);
        }

        public static AutomationElement GetElementByClassName(this AutomationElement wnd, string name)
        {
            return wnd.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.ClassNameProperty, name));
        }

        public static AutomationElementCollection GetElementsByClassName(this AutomationElement wnd, string name)
        {
            return wnd.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.ClassNameProperty, name));
        }

        public static AutomationElement GetElementByClassNameAndTitle(this AutomationElement wnd, string name, string title)
        {
            AutomationElementCollection btns = wnd.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.ClassNameProperty, name));

            foreach (AutomationElement item in btns)
            {
                if (item.Current.Name.Equal(title))
                {
                    return item;
                }
            }

            return null;
        }

        public static void SetValue(this AutomationElement wnd, string value)
        {
            ValuePattern valuePattern = (ValuePattern)wnd.GetCurrentPattern(ValuePattern.Pattern);
            valuePattern.SetValue(value);
        }

        public static void Click(this AutomationElement wnd)
        {
            var clickPattern = (InvokePattern)wnd.GetCurrentPattern(InvokePattern.Pattern);

            clickPattern.Invoke();
        }
    }
}
