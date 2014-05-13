//highlights de rij waar je muiscursor overheen gaat
function hlhr(Row, Action)
{
	//Row moet this zijn, Action moet 'over' of 'out' zijn
	var Cells = null;
	var newColor = null;
	Cells = Row.getElementsByTagName('td');
	if (Action == 'over') {
    	newColor = '#EBEBEB';
	}
	if (Action == 'out') {
		newColor = '';
	}
	var c = null;
    var rowCellsCnt	= Cells.length;
	for (c = 0; c < rowCellsCnt; c++) {
		Cells[c].setAttribute('bgcolor', newColor, 0);
	}
}