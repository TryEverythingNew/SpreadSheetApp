using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CptS322
{
    public abstract class Cell : INotifyPropertyChanged
    {
        private int RowInd;
        private int ColInd;
        protected string text;
        protected string realvalue = null;

        // get the event
        public event PropertyChangedEventHandler PropertyChanged;

        protected void NotifyPropertyChanged(String propertyName)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        // below is constructor and several properties
        public Cell(int rowindex, int colindex)
        {
            RowInd = rowindex;
            ColInd = colindex;
        }

        public int RowIndex
        {
            get { return RowInd; }
        }

        public int ColIndex
        {
            get { return ColInd; }
        }

        // in Cell class, when we set text differently, we invoke propertychanged
        public string Text
        {
            get { return text; }
            set
            {
                if (!String.Equals(text, value))
                {
                    text = value;
                    NotifyPropertyChanged("CellTextChange");
                }
            }
        }

        // only implement get property in Cell class
        public string Value
        {
            get
            {
                if (realvalue == null)
                    return null;
                return String.Copy(realvalue);
            }
        }
    }

    internal class Cell1 : Cell
    {
        private int RowInd;
        private int ColInd;

        public Cell1(int rowindex, int colindex) : base(rowindex, colindex)
        {
            RowInd = rowindex;
            ColInd = colindex;
        }

        // in inheritated class Cell1, we have both get and set, but only accessible within this package
        internal new string Value
        {
            get
            {
                if (realvalue == null)
                    return null;
                return realvalue;
            }
            set
            {
                if (!String.Equals(realvalue, value))
                {
                    realvalue = value;
                }
            }
        }
    }

    public class Spreadsheet
    {
        public Cell[][] cell_arr;
        public event EventHandler CellPropertyChanged;

        // initialize and subscriber the events
        public Spreadsheet(int row, int col)
        {
            cell_arr = new Cell1[row][];
            for (int i = 0; i < row; i++)
            {
                cell_arr[i] = new Cell1[col];
                for (int j = 0; j < col; j++)
                {
                    cell_arr[i][j] = new Cell1(i, j);
                    cell_arr[i][j].PropertyChanged += new PropertyChangedEventHandler(CellPropertyChange);
                }
            }

            
        }

        // update cells' values and it depends on whether we have '='
        void CellPropertyChange(object sender, PropertyChangedEventArgs e)
        {
            Cell1 cell = sender as Cell1;
            if (e.PropertyName == "CellTextChange")
            {
                if(cell.Text[0] == '=')
                {
                    int j = cell.Text[1] - 'A';
                    string tmp = cell.Text.Substring(2,cell.Text.Length-2);
                    int i = int.Parse(tmp) - 1;
                    cell.Value = cell_arr[i][j].Value;
                }
                else
                {
                    cell.Value = cell.Text;
                }
                CellPropertyChanged(sender as Cell, new EventArgs());
            }
        }

        public Cell GetCell(int row, int col)
        {
            if (cell_arr == null || row >= RowCount || col >= ColumnCount)
                return null;
            return cell_arr[row][col] as Cell;
        }

        public int RowCount
        {
            get {
                if (cell_arr != null)
                    return cell_arr.Length;
                else
                    return 0;
                }
        }

        public int ColumnCount
        {
            get
            {
                if (RowCount == 0 || cell_arr[0] == null)
                    return 0;
                else
                    return cell_arr[0].Length;
            }
        }
    }

    // abstract base class for expression tree node
    public abstract class ExpTreeNode
    {
        
    }

    // const node with only double value
    public class ConstNode : ExpTreeNode
    {
        public double value = double.NaN;

        public ConstNode(double v)
        {
            value = v;
        }
        
    }

    // operator node with operator as char and two nodes
    public class OpNode : ExpTreeNode
    {
        public char op = '\0';
        public ExpTreeNode left = null;
        public ExpTreeNode right = null;

        public OpNode(char o, ExpTreeNode l, ExpTreeNode r)
        {
            op = o;
            left = l;
            right = r;
        }

        
    }

    // set the variable node with a string as member
    public class VarNode : ExpTreeNode
    {
        public string Nodename = null;

        public VarNode(string name)
        {
            Nodename = name;
        }

        
    }

    // expression tree for store an expression of string type and computing its double value
    public class ExpTree
    {
        // dictionary to stor variables
        private Dictionary<string, double> vars = new Dictionary<string, double>();
        public ExpTreeNode root = null; // root of the tree

        public ExpTree(string expression)
        {
            vars = new Dictionary<string, double>(); // need to clear out the old dictionary
            root = Compile( expression);    // store the tree

        }

        // store the tree
        private ExpTreeNode Compile(string expression)
        {
            int index = GetOpIndex(expression); // get the operator index in the expression
            if (index == -1)
            {
                return BuildSimpleNode(expression); // if it's an expression without any operator, just build it.
            }
            else  // embeded calling to build node for operator
            {
                ExpTreeNode left = Compile(expression.Substring(0, index));
                ExpTreeNode right = Compile(expression.Substring(index + 1));
                return (new OpNode(expression[index], left, right));
            }
        }

        // built the node for const or variable
        private ExpTreeNode BuildSimpleNode(string expression)
        {
            double tmp;
            if(double.TryParse(expression, out tmp))    // can parse, means its const
            {
                return new ConstNode(tmp);
            }
            else    // cannot parse, its variable
            {
                return new VarNode(expression);
            }
        }

        public int GetOpIndex(string exp)   // loop throught expression to get the first operator to return default value -1
        {
            char tmp;
            for(int i = exp.Length - 1; i > 0; i--)
            {
                tmp = exp[i];
                if ( tmp == '+' || tmp == '-' || tmp == '*' || tmp == '/')
                {
                    return i;
                }
            }
            return -1; // means that the operator doesn't exist
        }

        // set a varaible node
        public void SetVar(string varName, double varValue)  
        {
            if (vars.ContainsKey(varName))  // if it's already in the dictionary, update its value
            {
                vars[varName] = varValue;
            }
            else    // if it's not in the dictionary, add it
            {
                vars.Add(varName, varValue);
            }
        }

        public double Eval()    // for evaluating the tree, just eval the node
        {
            return Eval(root);
        }

        public double Eval(ExpTreeNode node) // in-order traverse to see the node's numeric value
        {
            if (node == null)
            {
                return double.NaN;  // if null, return not a number
            }
            else if(node is ConstNode)  // if const, return its value
            {
                return (node as ConstNode).value;
            }
            else if(node is VarNode)        // if variable node, check the dictionary and return
            {
                if (vars.ContainsKey((node as VarNode).Nodename))
                    return vars[(node as VarNode).Nodename];
                else
                    return double.NaN;
            }
            else    // if operator node, embeded in order call to return the value of operation 
            {
                OpNode tmp = node as OpNode;
                char op = tmp.op; 
                switch (op)
                {
                    case '+': return (Eval(tmp.left) + Eval(tmp.right));
                    case '-': return (Eval(tmp.left) - Eval(tmp.right));
                    case '*': return (Eval(tmp.left) * Eval(tmp.right));
                    case '/': return (Eval(tmp.left) / Eval(tmp.right));
                }
                return double.NaN;  // if not +-/*, return not a number
            }
        }
    }
    
}
