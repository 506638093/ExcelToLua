/* 
 * ==============================================================================
 * Filename: 
 * Created:  2021 / 8 / 12 15:51
 * Author: HuaHua
 * Purpose: Excel to Lua
 * ==============================================================================
**/
using System;

class Program
{
    static void Main(string[] args)
    {
        if(args.Length != 2)
        {
            Console.WriteLine("The first parameter is the source file, and the second parameter is the output path");
            return;
        }

        Exporter.Export(args[0], args[1]);
    }
}
