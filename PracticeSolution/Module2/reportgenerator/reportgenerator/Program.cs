using System;

namespace reportgenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            Generator g = new Generator();
            g.initDataObject(/* jsonBody */);
            g.CreatePackage(args[0]);
            Console.ReadLine();
        }
    }
}
