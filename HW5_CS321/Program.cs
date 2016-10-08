using CptS322;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CptS322
{
    class Program
    {
        static void Main(string[] args)
        {
            string expression = "A1-B1-C1"; // set up default expreesion and the tree
            ExpTree tree = new ExpTree(expression);
            while (true)
            {
                // show the menu
                Console.Out.WriteLine("Menu (current expression is " + expression + ")");
                Console.Out.WriteLine("\t1 = Enter a new expression\n\t2 = Set a variable value");
                Console.Out.WriteLine("\t3 = Evaluate tree\n\t4 = Quit");

                // readline method to get the choice
                int choice;
                while (true)    // keep reading a number until it's between 1 and 4 
                {
                    string s = Console.ReadLine();
                    if(int.TryParse(s, out choice) && choice <=4 && choice > 0)
                    {
                        break;
                    }
                    // print out it's not a valid choice
                    Console.Out.WriteLine("Please enter a number between 1 to 4 or the app will ignore:");
                }
                if(choice == 4) // choice 4, finish the program
                {
                    Console.Out.WriteLine("Done");
                    break;
                }
                switch (choice)
                {
                    case 1: // choice 1, enter expression and setup the tree.
                        Console.Out.Write("Enter new expression: ");
                        expression = Console.ReadLine();
                        tree = new ExpTree(expression);
                        break;

                    case 2: // choice 2, readline to get variable name and variable value
                        Console.Out.Write("Enter variable name: ");
                        string var = Console.ReadLine();
                        Console.Out.Write("Enter variable value: ");
                        string value = Console.ReadLine();
                        double v = 0;
                        // check if value is double, if so, we add to dictionary, if not, tell the user
                        if (double.TryParse(value, out v))
                        {
                            tree.SetVar(var, v);
                        }
                        else
                        {
                            Console.Out.Write("Variable value is not valid double type: ");
                        }
                        break;

                    case 3: // print out the evaluated double value of the expression
                        Console.Out.WriteLine(tree.Eval());
                        break;
                }
            }
        }
    }
}
