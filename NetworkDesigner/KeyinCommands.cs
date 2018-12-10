/*--------------------------------------------------------------------------------------+
|   KeyinCommands.cs
|
+--------------------------------------------------------------------------------------*/

#region System Namespace
using System;
#endregion

namespace NetworkDesigner
{
    internal class KeyinCommands
    {
        public static void CommandKeyin(string unparsed)
        {
            //insert code here
        }

        public static void LocateCommand1Keyin(string unparsed)
        {
            LocateCommand1 tool = new LocateCommand1();
            tool.StartLocateCommand(unparsed);
        }

        public static void PlacementCommand1Keyin(string unparsed)
        {
            PlacementCommand1 tool = new PlacementCommand1();
            tool.StartPlacementCommand(unparsed);
        }

        //public static void Form2Keyin(string unparsed)
        //{
        //    Form2 form = new Form2();
        //    form.ShowForm(unparsed);
        //}

        public static void Form1Keyin(string unparsed)
        {
            Form1 form = new Form1();
            form.ShowForm(unparsed);
        }

        public static void Form2Keyin(string unparsed)
        {
            Form2 form = new Form2();
            form.ShowForm(unparsed);
        }
    }
}