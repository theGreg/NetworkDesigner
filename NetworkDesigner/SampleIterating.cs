using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetworkDesigner.TreeNode
{
    class SampleIterating
    {
        static void MainTest(string[] args)
        {
            TreeNode<MstnNodeData> treeRoot = NodeTestClass.GetSet1();
            foreach (TreeNode<MstnNodeData> node in treeRoot)
            {
                string indent = CreateIndent(node.Level);
                Console.WriteLine(indent + (node.Data.ToString()?? "null"));
            }
        }

        private static String CreateIndent(int depth)
        {
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < depth; i++)
            {
                sb.Append(' ');
            }
            return sb.ToString();
        }
    }
}
