using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;

namespace DemoTools
{
    class 线程池
    {
        static void Main1()
        {
            for (int i = 0; i < 10; i++)
            {
                ThreadPool.QueueUserWorkItem(m =>
                {
                    Console.WriteLine(Thread.CurrentThread.ManagedThreadId.ToString());
                });
            }
            Console.Read();
        }
    }
}
