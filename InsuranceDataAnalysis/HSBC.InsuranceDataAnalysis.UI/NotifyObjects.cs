////===============================================================================
//
//  Copyright © 2018 中软国际.HSBC业务线.第二事业部.保险与卡交付部 All rights reserved    
//  
//  Filename :NotifyObjects
//  Description :
//
//  Created by Tina at  2/2/2018 6:02:17 PM
//
////===============================================================================
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HSBC.InsuranceDataAnalysis.UI
{
    public class NotifyObjects : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        public void RaisePropertyChanged(string propertyName)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }
    }
}
