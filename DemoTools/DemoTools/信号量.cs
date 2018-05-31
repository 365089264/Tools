using System;
using System.Threading;
using System.Threading.Tasks;

namespace DemoTools
{
    class 信号量
    {
        static SemaphoreSlim semLim = new SemaphoreSlim(3); //3表示最多只能有三个线程同时访问
        static void Main2(string[] args)
        {
            for (int i = 0; i < 10; i++)
            {
                new Thread(SemaphoreTest).Start();
            }
            Console.Read();
        }
        static void SemaphoreTest()
        {
            semLim.Wait();
            Console.WriteLine("线程" + Thread.CurrentThread.ManagedThreadId.ToString() + "开始执行");
            Thread.Sleep(2000);
            Console.WriteLine("线程" + Thread.CurrentThread.ManagedThreadId.ToString() + "执行完毕");
            semLim.Release();
        }

        static void Main3(string[] args)
        {
            Console.WriteLine("主线程启动");
            //Task.Run启动一个线程
            //Task启动的是后台线程，要在主线程中等待后台线程执行完毕，可以调用Wait方法
            //Task task = Task.Factory.StartNew(() => { Thread.Sleep(1500); Console.WriteLine("task启动"); });
            Task task = Task.Run(() =>
            {
                Thread.Sleep(1500);
                Console.WriteLine("task启动");
            });
            Thread.Sleep(300);
            task.Wait();
            Console.WriteLine("主线程结束");
        }

        static void Main(string[] args)
        {
            for (int i = 0; i < 5; i++)
            {
                new Thread(Run1).Start();
            }
            for (int i = 0; i < 5; i++)
            {
                Task.Run(() => { Run2(); });
            }
            Console.Read();
        }
        static void Run1()
        {
            Console.WriteLine("Thread Id =" + Thread.CurrentThread.ManagedThreadId);
        }
        static void Run2()
        {
            Console.WriteLine("Task调用的Thread Id =" + Thread.CurrentThread.ManagedThreadId);
        }
    }
}
