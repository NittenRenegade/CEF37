using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Dynamic;
using System.Reflection;
using System.Runtime.InteropServices;

namespace ТестВК
{

    class ДинамикГК : DynamicObject, IDisposable
    {
        dynamic ГК1С;
        Type ТипГК;
        object App1C;

        void УничтожитьОбъект(object Объект)
        {
            Marshal.Release(Marshal.GetIDispatchForObject(Объект));
            Marshal.ReleaseComObject(Объект);


        }
        public ДинамикГК(dynamic ГК1С)
        {
            ТипГК = ГК1С.GetType();
            App1C = ГК1С.AppDispatch;
            this.ГК1С = ГК1С;


        }
        // установка свойства
        public override bool TrySetMember(SetMemberBinder binder, object value)
        {
            try
            {
                ТипГК.InvokeMember(binder.Name, BindingFlags.SetProperty, null, App1C, new object[] { value });
                return true;
            }
            catch (Exception)
            {
            }
            return false;
        }
        // получение свойства
        public override bool TryGetMember(GetMemberBinder binder, out object result)
        {
            try
            {
                result = ТипГК.InvokeMember(binder.Name, BindingFlags.GetProperty, null, App1C, null);
                return true;
            }
            catch (Exception)
            {
            }
            result = null;
            return false;
        }
        // вызов метода
        public override bool TryInvokeMember(InvokeMemberBinder binder, object[] args, out object result)
        {
            // dynamic method = members[binder.Name];
            // result = method((int)args[0]);
            // return result != null;
            if (binder.Name == "ЗакрытьОбъект")
            {

                Dispose();
                result = null;
                return true;

            }


            try
            {
                if (args.Length == 1 && args[0].GetType() == typeof(System.Object[]))
                    result = ТипГК.InvokeMember(binder.Name, BindingFlags.Static | BindingFlags.Instance | BindingFlags.Public | BindingFlags.InvokeMethod, null, App1C, (System.Object[])args[0]);
                else
                    result = ТипГК.InvokeMember(binder.Name, BindingFlags.Static | BindingFlags.Instance | BindingFlags.Public | BindingFlags.InvokeMethod, null, App1C, args);

                return true;
            }
            catch (Exception)
            {
            }
            result = null;
            return false;
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
                if (ГК1С != null)
                {
                    УничтожитьОбъект(ГК1С);
                    УничтожитьОбъект(App1C);
                    ГК1С = null;
                    App1C = null;
                    ТипГК = null;
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
