using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using ScriptHelp.Scripts;
using ScriptHelp.Scripts.Syntax;

namespace ScriptHelp.TaskPane
{
    /// <summary>
    /// Script TaskPane
    /// </summary>
    public partial class Script : UserControl
    {
        /// <summary>
        /// Initialize the controls in the object
        /// </summary>
        public Script()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Highlight words from KeyWords.
        /// </summary>
        private void UpdateSqlSyntax()
        {
            try
            {
                var syntaxHighlighter = new SyntaxHighlighter(this.txtScript);
                // comment
                syntaxHighlighter.AddPattern(new PatternDefinition(@"--[^\r\n]*"), new SyntaxStyle(Color.Green));
                // comment
                syntaxHighlighter.AddPattern(new PatternDefinition(@"(/\*([^*]|[\r\n]|(\*+([^*/]|[\r\n])))*\*+/)|(//.*)"), new SyntaxStyle(Color.Green));
                // zero strings
                syntaxHighlighter.AddPattern(new PatternDefinition("''"), new SyntaxStyle(Color.Red));
                // single quote strings
                syntaxHighlighter.AddPattern(new PatternDefinition(@"\'([^']|\'\')+\'"), new SyntaxStyle(Color.Red));
				// in brackets
				//syntaxHighlighter.AddPattern(new PatternDefinition(@"\[([^']|\'\')+\]"), new SyntaxStyle(Color.Black));  //was messing up the color for CREATE table statement, not sure if I need this
				// keywords1
				syntaxHighlighter.AddPattern(new PatternDefinition("XACT_ABORT", "BEGIN", "DELETE", "ROLLBACK", "COMMIT", "CREATE", "TABLE", "TRAN", "TRANSACTION", "OUTPUT", "USING", "BY", "TARGET", "WITH", "AS", "VALUES", "MERGE", "ON", "WHEN", "THEN", "UNION", "UPDATE", "SET", "WHERE", "GO", "APPEND", "INSERT", "INTO", "TRUNCATE", "REMOVE", "SELECT", "FROM", "TYPE", "FOLDER", "CABINET", "ORDER BY", "DESC", "ASC", "GROUP BY", "ALTER", "ADD", "DROP", "GROUP", "PRIMARY", "KEY", "IDENTITY", "IF"), new SyntaxStyle(Color.Blue));
                // keywords2
                syntaxHighlighter.AddPattern(new PatternDefinition("OBJECTS", "objects", "SYS", "sys"), new SyntaxStyle(Color.Green));
                // functions
                syntaxHighlighter.AddPattern(new PatternDefinition("$action", "object_id", "OBJECT_ID", "UPPER", "LOWER", "SUBSTR", "COUNT", "MIN", "MAX", "AVG", "SUM", "DATEDIFF", "DATEADD", "DATEFLOOR", "DATETOSTRING", "ID", "max", "MFILE_URL"), new SyntaxStyle(Color.Fuchsia));
                // operators
                syntaxHighlighter.AddPattern(new PatternDefinition("SOURCE", "MATCHED", "+", "-", ">", "<", "&", "|", "*", "**", "!", "=", "AND", "OR", "SOME", "ALL", "ANY", "LIKE", "NOT", "NULL", "NULLDATE", "NULLSTRING", "NULLINT", "IN", "EXISTS"), new SyntaxStyle(Color.Gray));
            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);

            }
        }

        /// <summary>
        /// Highlight words from KeyWords.
        /// </summary>
        private void UpdateDqlSyntax()
        {
            try
            {
                var syntaxHighlighter = new SyntaxHighlighter(this.txtScript);
                // zero strings
                syntaxHighlighter.AddPattern(new PatternDefinition("''"), new SyntaxStyle(Color.Red));
                // single quote strings
                syntaxHighlighter.AddPattern(new PatternDefinition(@"\'([^']|\'\')+\'"), new SyntaxStyle(Color.Red));
                // keywords1
                syntaxHighlighter.AddPattern(new PatternDefinition("UNION", "UPDATE", "SET", "WHERE", "GO", "APPEND", "INSERT", "INTO", "TRUNCATE", "REMOVE", "SELECT", "FROM", "TYPE", "FOLDER", "CABINET", "ORDER BY", "DESC", "ASC", "GROUP BY", "ALTER", "ADD", "DROP", "GROUP"), new SyntaxStyle(Color.Blue));
                // keywords2
                syntaxHighlighter.AddPattern(new PatternDefinition("OBJECTS", "objects"), new SyntaxStyle(Color.Green));
                // functions
                syntaxHighlighter.AddPattern(new PatternDefinition("UPPER", "LOWER", "SUBSTR", "COUNT", "MIN", "MAX", "AVG", "SUM", "DATEDIFF", "DATEADD", "DATEFLOOR", "DATETOSTRING", "ID", "MFILE_URL"), new SyntaxStyle(Color.Fuchsia));
                // operators
                syntaxHighlighter.AddPattern(new PatternDefinition("+", "-", ">", "<", "&", "|", "*", "**", "!", "=", "AND", "OR", "SOME", "ALL", "ANY", "LIKE", "NOT", "NULL", "NULLDATE", "NULLSTRING", "NULLINT", "IN", "EXISTS"), new SyntaxStyle(Color.Gray));

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);

            }
        }

        /// <summary>
        /// Update the script text with syntax formatting
        /// </summary>
        /// <param name="sender">contains the sender of the event, so if you had one method bound to multiple controls, you can distinguish them.</param>
        /// <param name="e">refers to the event arguments for the used event, they usually come in the form of properties/functions/methods that get to be available on it.</param>
        private void Script_Load(object sender, EventArgs e)
        {
            try
            {
                switch (Ribbon.AppVariables.FileType)
                {
                    case "SQL":
                        UpdateSqlSyntax();
                        break;
                    case "DQL":
                        UpdateDqlSyntax();
                        break;
					case "TXT":
						UpdateSqlSyntax();
						break;
					case "XML":
						UpdateSqlSyntax();
						break;
				}
                txtScript.Text = Ribbon.AppVariables.ScriptRange;

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);

            }
        }

        /// <summary>
        /// Copy the script text
        /// </summary>
        /// <param name="sender">contains the sender of the event, so if you had one method bound to multiple controls, you can distinguish them.</param>
        /// <param name="e">refers to the event arguments for the used event, they usually come in the form of properties/functions/methods that get to be available on it.</param>
        private void btnCopy_Click(object sender, EventArgs e)
        {
            try
            {
                this.txtScript.SelectAll();
                this.txtScript.Copy();
            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);

            }
        }

        /// <summary>
        /// Save the script text to a file
        /// </summary>
        /// <param name="sender">contains the sender of the event, so if you had one method bound to multiple controls, you can distinguish them.</param>
        /// <param name="e">refers to the event arguments for the used event, they usually come in the form of properties/functions/methods that get to be available on it.</param>
        private void btnSave_Click(object sender, EventArgs e)
        {
            try
            {
                SaveFileDialog s = new SaveFileDialog();
                switch (Ribbon.AppVariables.FileType)
                {
                    case "SQL":
                        s.FileName = "Update_" + Properties.Settings.Default.Sheet_Column_Table_Alias + ".sql";
                        s.Filter = "Structured Query Language | *.sql";
                        break;
                    case "DQL":
                        s.FileName = "Update_" + Ribbon.AppVariables.FirstColumnName + ".dql";
                        s.Filter = "Documentum Query Language | *.dql";
                        break;
					case "TXT":
						s.FileName = Properties.Settings.Default.Sheet_Column_Table_Alias + ".txt";
						s.Filter = "Text File | *.txt";
						break;
					case "XML":
						s.FileName = Properties.Settings.Default.Sheet_Column_Table_Alias + ".xml";
						s.Filter = "Extensible Markup Language | *.xml";
						break;
				}
                if (s.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    using (System.IO.StreamWriter sw = new System.IO.StreamWriter(s.FileName))
                    {
                        foreach (string line in txtScript.Lines)
                        {
                            sw.WriteLine(line);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);

            }
        }

    }
}
