using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ComInterop.DotNetComLibrary
{
    //    [Guid("4CD8F8D4-7470-4FD3-8BC2-627CE9B83D2D")]
    //    [Guid("0B98FF76-2883-4806-9030-B9D694626969")]
    //    [Guid("ADA7FAB6-4E8D-4D14-AC9E-10EAD91492B8")]


    [Guid("C7A0C50F-BCF8-4290-A3E3-8F8DA308FBFD")]
    [ComVisible(true)]
    public interface IServerObject
    {
        void ShowMessage(string message);
        int AddARandomNumber(int number);
        void ThrowAnException();
        void RegisterEventSink(IEventSink _eventSink);
    }

    [Guid("ADA7FAB6-4E8D-4D14-AC9E-10EAD91492B8")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class ServerObject : IServerObject
    {
        Random randomNumberGenerator = new Random();
        int eventCount = 0;
        System.Threading.Timer timer;

        public ServerObject()
        {
            timer = new System.Threading.Timer((state) => {
                eventSink?.OnEvent(eventCount++);
            }, null, TimeSpan.Zero, TimeSpan.FromMilliseconds(500));
        }

        public int AddARandomNumber(int number)
        {
            return number + randomNumberGenerator.Next(100);
        }

        public void ShowMessage(string message)
        {
            MessageBox.Show(message);
        }

        public void ThrowAnException()
        {
            throw new Exception("An exception");
        }

        IEventSink eventSink;

        public void RegisterEventSink(IEventSink _eventSink)
        {
            eventSink = _eventSink;
        }
    }

    [Guid("4CD8F8D4-7470-4FD3-8BC2-627CE9B83D2D")]
    [ComVisible(true)]
    public interface IEventSink
    {
        void OnEvent(int i);
    }


    /*
     * Need to pass an object in to ServerObject
     * when something happens in server object (say a timed event)
     * it calls a method on our object
     * 
     * 
     * 
     */
}
