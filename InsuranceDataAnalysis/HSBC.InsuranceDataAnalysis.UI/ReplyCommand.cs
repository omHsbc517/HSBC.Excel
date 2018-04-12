////===============================================================================
//
//  Copyright © 2018 中软国际.HSBC业务线.第二事业部.保险与卡交付部 All rights reserved    
//  
//  Filename :RealyCommand
//  Description :
//
//  Created by Tina at  2/2/2018 6:03:30 PM
//
////===============================================================================
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace HSBC.InsuranceDataAnalysis.UI
{
    public class ReplyCommand : ICommand
    {
        private Func<object, bool> _canExecute;
        private Action _execute;

        public ReplyCommand(Action execute) : this(execute, null) { }

        public ReplyCommand(Action execute, Func<object, bool> canExecute)
        {
            _canExecute = canExecute;
            _execute = execute;
        }

        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        public bool CanExecute(object parameter)
        {
            return _canExecute == null ? true : _canExecute(parameter);
        }

        public void Execute(object parameter)
        {
            _execute();
        }
    }
}
