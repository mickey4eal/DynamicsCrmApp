using DynamicsCrmApp.Services.Abstract;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DynamicsCrmApp
{
    class Program
    {
        static void Main(string[] args)
        {
            const string Value = "Welcome to Demo Dynamics CRM Application!\n\nPlease Enter Start to Initiate Connection to Dynamics CRM and Create a New Account";
            Console.WriteLine(Value);

            string command;

            do
            {
                command = Console.ReadLine() ?? "";
                command = command.ToLower();

                if (command.Equals("start"))
                {
                    new DynamicsCrmService().CreateCrmAccount();
                }
                else if (command.Equals("exit") || command.Equals("ex"))
                {
                    Environment.Exit(0);
                }
            }
            while (command != "ex" || command != "exit");
        }
    }
}
