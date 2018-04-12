using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;

namespace HSBC.InsuranceDataAnalysis.Utils
{
    public class BackgroundConverter : IValueConverter
    {
        #region IValueConverter  
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if (string.IsNullOrWhiteSpace((string)value) )
            {
                return Brushes.White;
            }
            else
                return Brushes.Red;

            // below change full row
            ListViewItem item = (ListViewItem)value;
            
            ListView listView = ItemsControl.ItemsControlFromItemContainer(item) as ListView; // Use the ItemsControl.ItemsContainerFromItemContainer(item) to get the ItemsControl.. and cast  

            // Get the index of a ListViewItem  
            int index = listView.ItemContainerGenerator.IndexFromContainer(item); // this is a state-of-art way to get the index of an Item from a ItemsControl  
            dynamic obj = item.DataContext;
            if (index % 2 == 0)
            {
                if (!string.IsNullOrWhiteSpace(obj.DiffDetail))
                {
                    return Brushes.Tomato;
                }
                return Brushes.WhiteSmoke;// .LightBlue;
            }
            else
            {
                if (!string.IsNullOrWhiteSpace(obj.DiffDetail))
                {
                    return Brushes.Tomato;
                }
                return Brushes.WhiteSmoke;// .Beige;
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
        #endregion IValueConverter  
    }
}
