const X = require('xlsx');
const fs = require('fs');
const ws = X.utils.aoa_to_sheet([['A','B','C'],[1,2,3],[4,5,6]]);
// autofilter
ws['!autofilter'] = { ref: 'A1:C1' };
// try freeze via !freeze
try { ws['!freeze'] = { xSplit: 0, ySplit: 1, topLeftCell: 'A2', activePane: 'bottomLeft', state: 'frozen' }; } catch(e){ console.log('freeze set err', e.message); }
// also try views
ws['!views'] = [{ ySplit: 1, state: 'frozen', topLeftCell: 'A2', activePane: 'bottomLeft' }];
const wb = X.utils.book_new();
X.utils.book_append_sheet(wb, ws, 'S1');
X.writeFile(wb, 'out.xlsx');
console.log('written');
