using DynamicsCrmApp.Services.Abstract;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace DynamicsCrmApp
{
    class Program
    {
        static void Main(string[] args)
        {
            const string Value = "Dynamics CRM Test App!\n\nPress Start to Import Data from CRM.";
            Console.WriteLine(Value);

            string command;

            do
            {
                command = Console.ReadLine() ?? "";
                command = command.ToLower();

                if (command.Equals("start"))
                {
                    new DynamicsCrmService().ConnectToDynamicsCrmAndGetEntityCollection();
                }
                else if (command.Equals("exit") || command.Equals("ex"))
                {
                    Console.WriteLine("Thanks for using application.\n");
                    Thread.Sleep(1000);
                    Environment.Exit(0);
                }
            }
            while (command != "ex" || command != "exit");
        }
    }
}
