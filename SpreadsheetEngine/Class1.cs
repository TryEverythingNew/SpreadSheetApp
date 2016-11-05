//Name: Xueliang Sun ID: 11387859
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace CptS322
{
    // implement the IUndoRedoCmd interface, i add a string property to help people understand the Undo's text
    public interface IUndoRedoCmd
    {
        IUndoRedoCmd Exec();
        string UndoText { get; }
    }

    // a single command related to restore cell's property
    public class RestoreCell : IUndoRedoCmd
    {
        private string m_oldtext; // for restoring text
        private uint backcolor; // for restoring background color
        private Cell m_cell; // to store the cell 
        private string propertychange; // store the Undo's text for people to understand this undo command

        // two constructors, one for restoring backcolr, one for restoring text change, can be extended in future 
        public RestoreCell(Cell c, uint color)
        {
            m_cell = c;
            backcolor = color;
            propertychange = "BackColor Change";
        }

        public RestoreCell(Cell c, string text)
        {
            m_cell = c;
            m_oldtext = text;
            propertychange = "Text Change";
        }

        public string UndoText
        {
            get
            {
                return propertychange;
            }
        }

        // implement the exec interface method and return an Undoredo command
        public IUndoRedoCmd Exec()
        {
            RestoreCell reverse = null; // store the reverse command
            switch (propertychange)
            {
                case "BackColor Change":
                    reverse = new RestoreCell(m_cell, m_cell.BackColor);
                    m_cell.BackColor = backcolor;
                    break;
                case "Text Change":
                    reverse = new RestoreCell(m_cell, m_cell.Text);
                    m_cell.Text = m_oldtext;
                    break;
            }
            return reverse;
        }
    }

    // a collection of undoredo commands
    public class UndoRedoCollection : IUndoRedoCmd
    {
        // use list to store the commands, propertychange for storing undo text
        private List<IUndoRedoCmd> m_actions;
        private string propertychange;

        // two constructors, one for single command, one for a list of commands
        public UndoRedoCollection(IUndoRedoCmd cmd)
        {
            m_actions = new List<IUndoRedoCmd>();
            m_actions.Add(cmd);
            propertychange = cmd.UndoText;
        }

        public UndoRedoCollection(List<IUndoRedoCmd> cmd)
        {
            m_actions = cmd;
            // if list is not empty, we set the propertychange based on first command's undo text
            if (cmd.Count != 0)
                propertychange = cmd[0].UndoText + ((cmd.Count > 1) ? " and ... ": "");
            else
                propertychange = null;
        }

        // give user of UI ablitity to change the undo text, which is string propertychange
        public void SetPropertyChange(string change)
        {
            propertychange = change;
        }

        public string UndoText
        {
            get
            {
                return propertychange;
            }
        }

        // implement the exec interface method and return an Undoredo collection
        public IUndoRedoCmd Exec()
        {
            List<IUndoRedoCmd> reverse_actions = new List<IUndoRedoCmd>();
            // pay attention to the reverse order, I insert every reverse command at the begining of the list
            for(int i = 0; i < m_actions.Count; i++)
            {
                IUndoRedoCmd reverse = m_actions[i].Exec();
                reverse_actions.Insert(0, reverse);
            }
            return new UndoRedoCollection( reverse_actions);
        }
    }

    public abstract class Cell : INotifyPropertyChanged
    {
        private int RowInd;
        private int ColInd;
        protected string text;
        protected string realvalue = null;
        protected uint backcolor = 0xFFFFFFFF;

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

        public uint BackColor
        {
            get
            {
                return backcolor;
            }
            set
            {
                if (backcolor != value)
                {
                    backcolor = value;
                    NotifyPropertyChanged("CellBackColorChange");
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

        // set the cell to default value which is useful when loading a spreadsheet
        public void SetToDefaults()
        {
            Text = null;
            BackColor = 0xFFFFFFFF;
        }

        // check if cell property is default, it's useful for saving a spreadsheet
        public bool HasNonDefaults
        {
            get
            {
                return !string.IsNullOrEmpty(Text) || BackColor != 0xFFFFFFFF;
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
                    NotifyPropertyChanged("CellValueChange");
                }
            }
        }
    }

    // an interface for expression tree, and it's realized in spreadsheet class so that expression tree can know a variable node's value
    public interface GetVarValue 
    {
        double GetValue(string s, Cell c);
    }

    // main class for munipulating cells and operations of cells
    public class Spreadsheet : GetVarValue
    {
        public Cell[][] cell_arr;
        Dictionary<Cell1, List<Cell1>> refertable = new Dictionary<Cell1, List<Cell1>>();
        //public bool isupdate = false; // indicate whether we are modifying cells based on user inputs or we are just updating refering cells
        // if we are updating, then there's no need to insert cell or delete cell from refer table, this boolean value is just for algorithm optimization

        // two stacks for undo and redo collection
        Stack<UndoRedoCollection> Undos = new Stack<UndoRedoCollection>();
        Stack<UndoRedoCollection> Redos = new Stack<UndoRedoCollection>();

        public event PropertyChangedEventHandler CellPropertyChanged; // inform gridview to update cells

        // method for form to add an undoredocollection
        public void AddUndo(UndoRedoCollection coll)
        {
            Undos.Push(coll);
        }

        // get the message of top of collection Undos
        public string UndoTopMsg()
        {
            if (Undos.Count == 0) return null;
            else
            {
                return Undos.Peek().UndoText;
            }
        }

        // get the message of top of collection Redos
        public string RedoTopMsg()
        {
            if (Redos.Count == 0) return null;
            else
            {
                return Redos.Peek().UndoText;
            }
        }

        // execute the Undo command and push to redo
        public void Undo()
        {
            if (Undos.Count == 0) return;
            UndoRedoCollection tmp = Undos.Pop().Exec() as UndoRedoCollection;
            if (tmp != null)
            {
                Redos.Push(tmp);
            }
        }

        // execute the Redo command and push to undo
        public void Redo()
        {
            if (Redos.Count == 0) return;
            UndoRedoCollection tmp = Redos.Pop().Exec() as UndoRedoCollection;
            if (tmp != null)
            {
                Undos.Push(tmp);
            }
        }

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


        // method for saving the spreadsheet
        public void Save(String filename)
        {
            // check the file exists or not and give user a warning in console
            if (File.Exists(filename))
            {
                Console.Out.WriteLine("File already exists, so erase the old content");
            }
            FileStream fs = null;
            // use try-catch to ensure file is openned successfully
            try
            {
                fs = File.Create(filename);
            }catch(Exception e)
            {
                Console.Out.WriteLine("cannot save to the file due to" + e.ToString());
                return;
            }
            
            // using utf-8 to encode characters, in fact, utf-8 is the default encoding in xmltextwriter
            XmlTextWriter xml_w = new XmlTextWriter(fs, Encoding.UTF8);
            xml_w.WriteStartDocument();

            // set the indentation for human read
            xml_w.Formatting = Formatting.Indented;
            xml_w.Indentation = 4;

            // start spreadsheet element
            xml_w.WriteStartElement("spreadsheet");
            // traverse the cells and save cell properties with not default values
            for (int i = 0; i < RowCount; i++)
            {
                for (int j = 0; j < ColumnCount; j++)
                {
                    // if not default property, we save its row index, col index, and background color, and text string
                    if (cell_arr[i][j].HasNonDefaults)
                    {
                        xml_w.WriteStartElement("cell");
                        xml_w.WriteElementString("row", i.ToString());
                        xml_w.WriteElementString("col", j.ToString());
                        xml_w.WriteElementString("bg", cell_arr[i][j].BackColor.ToString());
                        xml_w.WriteElementString("text", cell_arr[i][j].Text);
                        xml_w.WriteEndElement();
                        
                    }
                }
            }
            xml_w.WriteEndElement();
            // write the end of document
            xml_w.WriteEndDocument();

            xml_w.Close();
            fs.Dispose();

        }

        // method for loading the spreadsheet
        public void Load(String filename)
        {
            // check the file exists or not, if not exists, return
            if (!File.Exists(filename))
            {
                Console.Out.WriteLine("File doen't exists, so return directly");
                return;
            }
            // use try-catch to ensure file is openned successfully
            FileStream fs = null;
            try
            {
                fs = File.OpenRead(filename);
            }
            catch (Exception e)
            {
                Console.Out.WriteLine("cannot load the file due to" + e.ToString());
                return;
            }

            XmlTextReader reader = null;

            // use try catch to avoid we read something invalid, and we don't want to crash
            try
            {
                // before we really load the cells, we need to clear out all the cells' properties
                for (int i = 0; i < RowCount; i++)
                {
                    for (int j = 0; j < ColumnCount; j++)
                    {
                        cell_arr[i][j].SetToDefaults();
                    }
                }
                // also, we need to clear the undo redo stack before loading
                Undos.Clear();
                Redos.Clear();

                // Load the reader with the data file and ignore all white space nodes.      
                reader = new XmlTextReader(fs);
                reader.WhitespaceHandling = WhitespaceHandling.None;

                //initialize some cell properties
                string prev = null, text = null; // prev is the previous element tag, text is to store cell text
                int row = -1, col = -1; // to store cell's row/col index
                bool row_read = false, col_read = false; // whether we read a row/col index successfully, false means fail, true means success
                uint color = 0; // to store color property
                while (reader.Read())
                {
                    switch (reader.NodeType) // switch the nodetype to find the suitable cell property
                    {
                        case XmlNodeType.Element:
                            prev = reader.Name; // set the previous element tag
                            break;
                        case XmlNodeType.Text:
                            Console.Write(reader.Value);
                            if(prev == "row")
                            {
                                // if previous tag is "row", then we can load the row index now
                                if(int.TryParse(reader.Value, out row) && row >= 0 && row < RowCount)
                                {
                                    row_read = true;
                                }
                                else
                                    row_read = false;
                            }
                            else if (prev == "col")
                            {
                                // if previous tag is "col", then we can load the col index now
                                if (int.TryParse(reader.Value, out col) && col >= 0 && col < ColumnCount)
                                {
                                    col_read = true;
                                }
                                else
                                    col_read = false;
                            }
                            else if (prev == "text")
                            {
                                // if previous tag is "text", then we can load the cell's text content now
                                text = reader.Value;
                                if (row_read && col_read) // only load if row/col index loaded successfully
                                    cell_arr[row][col].Text = text;
                            }
                            else if (prev == "bg")
                            {
                                // if previous tag is "bg", then we can load the cell's background color now
                                if (uint.TryParse(reader.Value, out color) && row_read && col_read)
                                { // only load if row/col index loaded successfully
                                    cell_arr[row][col].BackColor = color;
                                }
                            }
                            break;
                        default:
                            prev = null;
                            break;
                    }

                }
            }
            catch (Exception e)
            {
                Console.Out.WriteLine("cannot load the file using XMLTextReader due to" + e.ToString());
                return;
            }

            reader.Dispose();
            fs.Dispose();
        }


        // update cells' values and it depends on whether we have '='
        void CellPropertyChange(object sender, PropertyChangedEventArgs e)
        {
            Cell1 cell = sender as Cell1;
            if (e.PropertyName == "CellTextChange")
            {
                //every time we are modifying the cell's text, we need to delete the cell from reference table since it's just text
                DeleteCellfromRefertable(cell);
                
                if(cell.Text == null || cell.Text == string.Empty)
                {
                    cell.Value = null;
                }
                // if text starts with '=', we set up an expression tree to evaluate its value
                else if (cell.Text[0] == '=')
                {
                    ExpTree tree = new ExpTree(cell.Text.Substring(1), this, cell);
                    cell.Value = tree.Eval().ToString();
                }
                else
                { // if text doesn't start with '=', we deem it as general text, so we directly set the cell value to the text
                    cell.Value = cell.Text;
                }
                CellPropertyChanged(sender as Cell, e);
            }
            if (e.PropertyName == "CellValueChange") // when value is changed, we need to update referring cells
            {
                // once the cell is changed, we need to update all the cells that refer to this changed cell
                if (refertable.ContainsKey(cell) && refertable[cell].Count != 0)
                {
                    //isupdate = true; // now, we are updating referring cells, not modifying cells
                    for (int i = 0; i < refertable[cell].Count && (refertable[cell][i] != null); i++)
                    {
                        Cell1 cell_tmp = refertable[cell][i];
                        ExpTree tree_tmp = new ExpTree(cell_tmp.Text.Substring(1), this, cell_tmp);
                        cell_tmp.Value = tree_tmp.Eval().ToString();
                        CellPropertyChanged(cell_tmp as Cell, e);
                    }
                }
            }
            if (e.PropertyName == "CellBackColorChange") // when background is changed, we need to notify the gridview
            {
                CellPropertyChanged(sender as Cell, e);
            }
        }

        // every time the cell is modifying (not updating) its text property, we need to delete the cell from refer table
        void DeleteCellfromRefertable(Cell1 c)
        {
            foreach( KeyValuePair<Cell1, List<Cell1>> keyvaluepair in refertable)
            {
                if(keyvaluepair.Value != null)
                {
                    keyvaluepair.Value.Remove(c); // delete the cell from the dictionary, refertable
                }
            }
        }

        // refered is the cell that is refered, wanttorefer is the cell that needs the value from refered cell
        void InsertCellfromRefertable(Cell1 refered, Cell wanttorefer) 
        {
            if (!refertable.ContainsKey(refered))
            {
                refertable.Add(refered, new List<Cell1> { wanttorefer as Cell1 });
            }
            else if (refertable[refered] == null)
            {
                refertable[refered] = new List<Cell1> { wanttorefer as Cell1 };
            }
            else {
                List<Cell1> list = refertable[refered];
                for(int i = 0; i < list.Count; i++)
                {
                    if (Object.ReferenceEquals(list[i], wanttorefer as Cell1))
                        return;
                }
                refertable[refered].Add(wanttorefer as Cell1);
            }
        }

        public Cell GetCell(int row, int col)
        {
            if (cell_arr == null || row >= RowCount || col >= ColumnCount)
                return null;
            return cell_arr[row][col] as Cell;
        }

        // implement the interface for giving cell values to the expression tree's variable node
        public double GetValue(string s, Cell c)
        {
            int col_ind = s[0] - 'A';
            int row_ind = 0;
            double val = double.NaN;
            if (Int32.TryParse(s.Substring(1), out row_ind) && row_ind >= 1 && row_ind <= cell_arr.Length && col_ind >= 0 && col_ind < cell_arr[0].Length)
            {
                // if we are modifying the cell (not updating the cell), we want to register this cell into our refertable
                //if (!isupdate)
                //{
                Cell1 cell_tmp = cell_arr[row_ind - 1][col_ind] as Cell1;
                InsertCellfromRefertable(cell_tmp, c);
                //}

                string tmp = cell_arr[row_ind - 1][col_ind].Value;
                if (double.TryParse(tmp, out val))
                {
                    return val;
                }
            }
            return double.NaN;
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
        // dictionary to store variables
        GetVarValue dictionary;
        Cell owner; // owner owns this expression tree
        public ExpTreeNode root = null; // root of the tree

        public ExpTree(string expression, GetVarValue dic, Cell c)
        {
            dictionary = dic;
            owner = c;
            expression = expression.Replace(" ", string.Empty);
            root = Compile( expression);    // store the tree
        }

        // store the tree
        private ExpTreeNode Compile(string expression)
        {
            if (expression == null) // no expression at all
                return null;
            int index = GetOpIndex(expression); // get the operator index in the expression
            if (index == -1)
            {
                if(expression[0] == '(')    // if the expression starts with '(', we need to remove this parenthesis
                {
                    return Compile(expression.Substring(1, expression.Length - 2));
                }
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
                //if (!vars.ContainsKey(expression))  // only add the variable to the dictionary when it's not in there
                //{
                //    SetVar(expression, 0);
                //}
                return new VarNode(expression);
            }
        }

        public int GetOpIndex(string exp)   // loop through expression to get the first operator or return default value -1
        {
            char tmp;
            int parenthesis = 0;    // number of parenthesis
            int muldiv = -1; // index for * or /
            for (int i = exp.Length - 1; i >= 0; i--)
            {
                tmp = exp[i];
                if( tmp == ')')
                {
                    parenthesis++;
                    continue;
                }
                if( tmp == '(')
                {
                    parenthesis--;
                    if (parenthesis < 0)
                    {
                        throw new Exception(" left parenthesis more than right parenthesis");   // throw error if parenthesis number incorrect
                    }
                    continue;
                }
                if ( parenthesis == 0 && ( tmp == '+' || tmp == '-' || tmp == '*' || tmp == '/'))
                {
                    if(tmp == '+' || tmp == '-')    // if operator is + or -, we return directly
                    {
                        return i;
                    }
                    else
                    {
                        if (muldiv == -1)   // if operator is * or /, and we haven't set the index for it, we set muldiv index
                            muldiv = i;
                    }
                }
            }
            if (parenthesis != 0)
            {
                throw new Exception(" left parenthesis unequal to right parenthesis"); // throw error if parenthesis number incorrect
            }
            return muldiv; // means that we return the muldiv index since we didn't find + or -
        }

        // set a varaible node
        //public void SetVar(string varName, double varValue)  
        //{
        //    if (vars.ContainsKey(varName))  // if it's already in the dictionary, update its value
        //    {
        //        vars[varName] = varValue;
        //    }
        //    else    // if it's not in the dictionary, add it
        //    {
        //        vars.Add(varName, varValue);
        //    }
        //}

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
                return dictionary.GetValue((node as VarNode).Nodename, owner);
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
