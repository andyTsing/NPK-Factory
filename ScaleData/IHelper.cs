using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ScalesData
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual), Guid("5B45779A-4DD5-488F-8384-2E724CE5F1E2")]
    public interface IHelper
    {
        void open();
        void close();

        //void showView();
        void showView(bool autoConnect = false);
        void exportExcel(string startDate, string endDate, string password = "");

        void addScales(string tag);
        string[] getAllScales();
        //void insertScalesValue(string scaleTag, float weight);
        void insertScalesValue(string scaleTag, float weight, string datetime = null);
    }
}
