using System;
using System.Collections;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Diagnostics;
//using System.Windows.Forms;

namespace ТестВК
{
    [Guid("ab634004-f13d-11d0-a459-004095e1daea"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IAsyncEvent
    {
        void SetEventBufferDepth(int lDepth);
        void GetEventBufferDepth(ref int plDepth);
        void ExternalEvent(string bstrSource, string bstrMessage, string bstrData);
        void CleanBuffer();
    }


    [ComImport]
    [Guid("EFE19EA0-09E4-11D2-A601-008048DA00DE")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IExtWndsSupport
    {
        void GetAppMainFrame(
            [Out] out IntPtr hwnd);

        void GetAppMDIFrame(
            [Out] out IntPtr hwnd);

        //void CreateAddInWindow(
        //    [In] string bstrProgID,
        //    [In] string bstrWindowName,
        //    [In] int lStyles,
        //    [In] int lExStyles,
        //    [In, Out] ref System.Drawing.Size rctSize,
        //    [In] int Flags,
        //    [In, Out] ref IntPtr pHwnd,
        //    [In, Out, MarshalAs(UnmanagedType.IDispatch)] ref object pDisp);
    }

    public class ТестВК :IDisposable
{
    [DllImport("user32.dll")]
    public static extern int SetWindowLong(int hWnd, int nIndex, int dwNewLong);

    public static void SetOwner(int child, int owner)
    {
        SetWindowLong(child,
            -8, // GWL_HWNDPARENT
            owner);
    }

    public dynamic Object1C;
    public dynamic ГК;

public ТестВК(object Object1C)
    {
        this.Object1C = Object1C;
        ГК = new ДинамикГК(Object1C);

}

    public dynamic Новый( params object[] Параметры)
    {

        return ГК.NewObject(Параметры);
    }
 
public dynamic ПолучитьОписаниеТиповСтроки(int ДлинаСтроки)
{
    return Новый("ОписаниеТипов", "Строка", null, Новый("КвалификаторыСтроки", ДлинаСтроки, ГК.ДопустимаяДлина.Переменная));

} // ПолучитьОписаниеТиповСтроки()

// Служебная функция, предназначенная для получения описания типов числа, заданной разрядности.
// 
// Параметры:
//  Разрядность 			- число, разряд числа.
//  РазрядностьДробнойЧасти - число, разряд дробной части.
//
// Возвращаемое значение:
//  Объект "ОписаниеТипов" для числа указанной разрядности.
//
public dynamic ПолучитьОписаниеТиповЧисла(int Разрядность, int РазрядностьДробнойЧасти = 0, int ЗнакЧисла = 1)
{
    object КвалификаторЧисла;
    if (ЗнакЧисла == 1)
        КвалификаторЧисла = Новый("КвалификаторыЧисла", Разрядность, РазрядностьДробнойЧасти);
    else
        КвалификаторЧисла = Новый("КвалификаторыЧисла", Разрядность, РазрядностьДробнойЧасти, ЗнакЧисла);


    return Новый("ОписаниеТипов", "Число", КвалификаторЧисла);

} // ПолучитьОписаниеТиповЧисла()

public dynamic ПолучитьОписаниеТиповДаты(dynamic ЧастиДаты)
{
    return Новый("ОписаниеТипов", "Дата", null, null, Новый("КвалификаторыДаты", ЧастиДаты));

} // ПолучитьОписаниеТиповДаты()
public object ВернутьТз()
{

    dynamic Тз = Новый("ТаблицаЗначений");
    dynamic Колонки = Тз.Колонки;
    Колонки.Добавить("КолонкаЧисло", ПолучитьОписаниеТиповЧисла(9, 0));
    Колонки.Добавить("КолонкаЧисло4", ПолучитьОписаниеТиповЧисла(7, 2));
    Колонки.Добавить("КолонкаДата", ПолучитьОписаниеТиповДаты(ГК.ЧастиДаты.ДатаВремя));
    Колонки.Добавить("КолонкаСтрока", ПолучитьОписаниеТиповСтроки(0));
    dynamic Стр = Тз.Добавить();
    Стр.КолонкаЧисло = 4;
    Стр.КолонкаЧисло4 = 3.14m;
    Стр.КолонкаДата = DateTime.Now;
    Стр.КолонкаСтрока = "Заполнение из ВК";
    return Тз;
}
    void УничтожитьОбъект(object Объект)
{
    Marshal.Release(Marshal.GetIDispatchForObject(Объект));
    Marshal.ReleaseComObject(Объект);


}

    public void ВызватьГС()
    {

        GC.Collect();
        GC.WaitForPendingFinalizers();
    }

    //public string СоздатьОкно() {
    ////    m_1cApp.AppDispatch.Сообщить("Привет из ВК");
    //// не работает для упрощения работы создань динамический объект ДинамикГК
    //// выполняющий аналогичную фунцию
    
    //        IExtWndsSupport n;
    //        ГК.Сообщить("Привет из ВК", ГК.СтатусСообщения.Важное);
    //        n = (IExtWndsSupport)Object1C;
    //        IntPtr hwnd;
    //         n.GetAppMainFrame(out hwnd);

    //// Создаем форму, устанавливаем нативные хэндлы и устанвливаем окно 1С владельцем нетовского окна
    //        var form = new Form1();
    //        form.CreateControl();
    //        SetOwner(form.Handle.ToInt32(), hwnd.ToInt32());
    //        form.EventTo1C = Object1C as IAsyncEvent;
           
    //      form.Show();
            
    //        return "Методы ВК выполнены!";
    //}

    public void Закрыть()
    {
         Dispose();
         ВызватьГС();

    }

    bool disposed = false;
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    // Protected implementation of Dispose pattern. 
    protected virtual void Dispose(bool disposing)
    {
        if (disposed)
            return;

        if (disposing)
        {
            if (Object1C != null)
            {
                УничтожитьОбъект(Object1C);
                 Object1C = null;
                
                ((IDisposable) ГК).Dispose();
                    ГК = null;
                // Free any other managed objects here. 
                //
                }
        }

        // Free any unmanaged objects here. 
        //
        disposed = true;
    }
}
}
