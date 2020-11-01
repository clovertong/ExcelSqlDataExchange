using System;
using System.Windows.Input;

namespace ExcelSqlDataExchange.Models
{
    public class DelegateCommand : ICommand
    {
        Action _execute;
        Func<bool> _canExecute;
        public DelegateCommand(Action execute, Func<bool> canExecute)
        {
            this._execute = execute;
            this._canExecute = canExecute;
        }

        public DelegateCommand(Action execute)
        {
            this._execute = execute;
        }

        public bool CanExecute(object parameter)
        {
            if (_canExecute != null)
                return _canExecute();

            return true;
        }

        public event EventHandler CanExecuteChanged;

        public void Execute(object parameter)
        {
            this._execute();
        }
    }
}
