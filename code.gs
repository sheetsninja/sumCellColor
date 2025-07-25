
/**
* Get the sum of the cells matching the background of the cell with the formula.
* @param {range} sumRange Range to be evaluated
* @param {range} checkboxCell Toggle this checkbox to refresh sum
* @customfunction
*/
function SUMCELLS(sumRange, checkboxCell) {
    let activeRange = SpreadsheetApp.getActiveRange();
    let activeSheet = activeRange.getSheet();
    let color = activeRange.getBackground();
    let formula = activeRange.getFormula();
    let match = formula.match(/^\s*=\s*SUMCELLS\s*\(\s*([$A-Z]+\d+:[$A-Z]+\d+)(?:\s*,.*)?\s*\)$/i);
    let rangeA1Notation = match ? match[1] : null;
    let range = activeSheet.getRange(rangeA1Notation);
    let bg = range.getBackgrounds();
    let values = range.getValues();
    let sum = 0;

    for (i = 0; i < bg.length; i++)
        for (j = 0; j < bg[0].length; j++)
            if (bg[i][j] == color)
                sum = sum + values[i][j];
    return sum;
}
