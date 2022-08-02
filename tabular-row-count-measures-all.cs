// Author:
//     Tim Keeler
//     Created 05-16-2022

// Description:
//     Create a COUNTROWS measure for each table in the model with the option to create a corresponding 
//     calculation group. Script ignores creating measures for hidden date tables resulting from leaving 
//     Auto Date/Time enabled in Power BI, calculation groups, and whichever table is selected to store 
//     the measures. It is recommended to create a separate table for the purpose of storing DAX measures.


#r "Microsoft.VisualBasic"
using Microsoft.VisualBasic;
using System.Windows.Forms;

// Enter name of table where measures will be created
string measureTableName = 
    Interaction.InputBox(
    "Name of the table where measures will be created", 
    "Enter table name", "", 740, 400
    );

// Check if table exists in the model
if(!Model.Tables.Any(Table => string.Equals(Table.Name, measureTableName, StringComparison.InvariantCultureIgnoreCase)))
{
    Error(measureTableName + " does not exist in the model");
    return;
};

var measureTable = Model.Tables[measureTableName];

// Enter name of folder where measures will be stored 
string folderName = 
    Interaction.InputBox(
    "Name of the folder where the measures will be located", 
    "Enter folder name", "", 740, 400
    );

// Empty list that will be populated with new measures
List<string> measureList = new List<string>();
var answerCalcGroup = 0;
// Option to create a corresponding calculation group referencing new measures
switch(MessageBox.Show("Create calculation group with table measures?", "Calculation Group", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
{
    case DialogResult.Yes:
        answerCalcGroup = 1;
        string calcGroupName = 
            Interaction.InputBox(
            "Enter a name for the calculation group", 
            "Calculation group name", "", 740, 400
            );

        var newCalculationGroup = Model.AddCalculationGroup ();

        (Model.Tables["New Calculation Group"] as CalculationGroupTable).CalculationGroup.Precedence = 1;
        newCalculationGroup.Name = calcGroupName;
        newCalculationGroup.Columns["Name"].Name = "Table Name";

        foreach(var t in Model.Tables)
        {
            if(!(string.Equals(t.Name, measureTableName, StringComparison.InvariantCultureIgnoreCase)) && t.ObjectType != ObjectType.CalculationGroupTable && !t.Name.Contains("LocalDateTable") && !t.Name.Contains("DateTableTemplate"))
            {    
                var newMeasure = measureTable.AddMeasure(
                // Name of measure
                "Records (" + t.Name + ")",
                // DAX expression
                "COUNTROWS ( " + t.DaxObjectFullName + " )"
                );
                // Format string of measure
                newMeasure.FormatString = "#,#";
                // Measure description for additional context
                newMeasure.Description = "Count of records in " + t.DaxObjectFullName;
                // Add measure to display folder in table defined by [measureTable] variable
                newMeasure.DisplayFolder = folderName;
                // Create calculation item for each new measure created
                var newCalculationItem = newCalculationGroup.AddCalculationItem(newMeasure.Name);
                // Calculation item DAX expression references new measure
                newCalculationItem.Expression = newMeasure.DaxObjectFullName;
                measureList.Add(newMeasure.DaxObjectFullName);
            }
        };
        break;
    case DialogResult.No:
        foreach(var t in Model.Tables)
        {
            if(!(string.Equals(t.Name, measureTableName, StringComparison.InvariantCultureIgnoreCase)) && t.ObjectType != ObjectType.CalculationGroupTable && !t.Name.Contains("LocalDateTable") && !t.Name.Contains("DateTableTemplate"))
            {    
                var newMeasure = measureTable.AddMeasure(
                // Name of measure
                "Records (" + t.Name + ")",
                // DAX expression
                "COUNTROWS ( " + t.DaxObjectFullName + " )"
                );
                // Format string of measure
                newMeasure.FormatString = "#,#";
                // Measure description for additional context
                newMeasure.Description = "Count of records in " + t.DaxObjectFullName;
                // Add measure to display folder in table defined by [measureTable] variable
                newMeasure.DisplayFolder = folderName;
                measureList.Add(newMeasure.DaxObjectFullName);
            }
        };
        break;
};

var totalMeasure = measureTable.AddMeasure(
    "Total Records", 
    String.Join(" + ", measureList)
    );
    totalMeasure.FormatString = "#,#";
    totalMeasure.Description = "Count of records in all selected tables";
    totalMeasure.DisplayFolder = folderName;

var measureCount = Convert.ToString(measureList.Count + 1);

string msgTitle = "Script has completed";

if(answerCalcGroup > 0)
{
    string msgBody = measureCount + " measures and 1 calculation group created";
    MessageBox.Show(msgBody, msgTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
}
else
{
    string msgBody = measureCount + " measures created";
    MessageBox.Show(msgBody, msgTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
}
