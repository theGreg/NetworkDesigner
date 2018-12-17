using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetworkDesigner.TreeNode
{
    public class NodeTestClass
    {
        public static TreeNode<MstnNodeData> GetSet1()
        {
            TreeNode<MstnNodeData> root = new TreeNode<MstnNodeData>(new MstnNodeData((long) 0));
            {
                TreeNode<MstnNodeData> node0 = root.AddChild(new MstnNodeData(2538));
                TreeNode<MstnNodeData> node1 = root.AddChild(new MstnNodeData(5844));
                TreeNode<MstnNodeData> node2 = root.AddChild(new MstnNodeData(5558));
                {
                    TreeNode<MstnNodeData> node20 = node2.AddChild(null);
                    TreeNode<MstnNodeData> node21 = node2.AddChild(new MstnNodeData(556));
                    {
                        TreeNode<MstnNodeData> node210 = node21.AddChild(new MstnNodeData(220));
                        TreeNode<MstnNodeData> node211 = node21.AddChild(new MstnNodeData(548));
                    }
                }
                TreeNode<MstnNodeData> node3 = root.AddChild(new MstnNodeData(878));
                {
                    TreeNode<MstnNodeData> node30 = node3.AddChild(new MstnNodeData(89866));
                }
            }

            return root;
        }
    }
}
