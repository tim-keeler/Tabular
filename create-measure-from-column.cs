// Author: Tim Keeler
// Created: 09-30-2022

// Prompt user to choose a DAX expression and create a measure for each selected column based on that choice.
// Measure will be created in the same table where the selected column is from.

// Credit to Stephen Maguire for his custom userform function: https://gist.github.com/samaguire/925a261ab5fcb86c944196b4c316e116



#r "Microsoft.VisualBasic"
using Microsoft.VisualBasic;
using System.Windows.Forms;

var promptForVariables = true;                                      // Set to true to prompt user for defaultMeasureExpressionName
var defaultMeasureExpressionName = "SUM";                           // Default combobox value when initialized
var defaultMeasureExpression = @"{0} ( {1} )";                      // Default DAX measure expression
var defaultMeasureDescription =                                     // Default DAX measure description
    @"This measure is the {0} of column [{1}] from table [{2}]";

// Create dictionary with key/value pairs to pass DAX expression and string used in measure name & description
Dictionary<string, string> expressionDict = new Dictionary<string, string>()
{
    {"AVERAGE", "Average"},
    {"DISCTINCTCOUNT", "Distinct Count"},
    {"MAX", "Max"},
    {"MEDIAN", "Median"},
    {"MIN", "Min"},
    {"SUM", "Sum"}
};

// Custom userform function
Func<string, string, string, string> ComboBox = (string promptText, string titleText, string defaultText) =>
{
    var labelText = new Label()
    {
        Text = promptText,
        Dock = DockStyle.Fill,
    };

    var comboboxText = new ComboBox()
    {
        Text = defaultText,
        Dock = DockStyle.Bottom,
    };
    // Populate combobox with dictionary keys
    foreach(var x in expressionDict)
    {
        comboboxText.Items.Add(x.Key);
    }

    var panelButtons = new Panel()
    {
        Height = 30,
        Dock = DockStyle.Bottom
    };
    
    var buttonOK = new Button()
    {
        Text = "OK",
        DialogResult = DialogResult.OK,
        Top = 8,
        Left = 120
    };

    var buttonCancel = new Button()
    {
        Text = "Cancel",
        DialogResult = DialogResult.Cancel,
        Top = 8,
        Left = 204
    };

    var formComboBox = new Form()
    {
        Text = titleText,
        Height = 143,
        Padding = new System.Windows.Forms.Padding(8),
        FormBorderStyle = FormBorderStyle.FixedDialog,
        MinimizeBox = false,
        MaximizeBox = false,
        StartPosition = FormStartPosition.CenterScreen,
        AcceptButton = buttonOK,
        CancelButton = buttonCancel
    };

    formComboBox.Controls.AddRange(new Control[] { labelText, comboboxText, panelButtons });
    panelButtons.Controls.AddRange(new Control[] { buttonOK, buttonCancel });

    return formComboBox.ShowDialog() == DialogResult.OK ? comboboxText.Text : null;
};

// Make sure user has selected at least one column
if (!Selected.Columns.Any())
{
    ScriptHelper.Error("No column(s) selected.");
    return;
}

// Prompt for input variables
if (promptForVariables)
{
    defaultMeasureExpressionName = ComboBox(
        "Select the type of calculation that will be used to create a DAX measure based on the selected column(s)",
        "Select Calculation Type",
        defaultMeasureExpressionName
        );
    if (defaultMeasureExpressionName == null)
    {
        return;
    }
}

// Find user selection in dictionary
foreach(KeyValuePair<string, string> x in expressionDict)
{
    if(x.Key == defaultMeasureExpressionName)
    {
        int measureCount = 0;                       // Count number of measures created
        var MeasureExpressionName = @"{0} of {1}";  // Default measure name string

        foreach(var c in Selected.Columns)          // Create measure for each selected column
        {
            var newMeasure = c.Table.AddMeasure(
            string.Format(MeasureExpressionName, x.Value, c.Name),                  // Name of measure     
            string.Format(defaultMeasureExpression, x.Key, c.DaxObjectFullName),    // DAX expression
            c.DisplayFolder                                                         // Display folder
            );
            newMeasure.FormatString = c.FormatString;                               // Inherit measure formatting from column
            newMeasure.Description = string.Format(defaultMeasureDescription, x.Value.ToLower(), c.Name, c.Table.Name);

            measureCount += 1;
        }
        
        // Display message box with number and type of measures created
        var msgboxMessage = @"({0}) {1} {2} been created.";
        string msgboxTitle = "Script Complete";

        if(measureCount == 1)
        {
            string msgboxSuffix = "measure has";
            MessageBox.Show(
                string.Format(msgboxMessage, measureCount, defaultMeasureExpressionName, msgboxSuffix),
                msgboxTitle,
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            );
        }
        else
        {
            string msgboxSuffix = "measures have";
            MessageBox.Show(
                string.Format(msgboxMessage, measureCount, defaultMeasureExpressionName, msgboxSuffix),
                msgboxTitle,
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            );
        }
    }
}


