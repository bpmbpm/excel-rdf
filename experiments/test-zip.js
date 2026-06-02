const X = require('xlsx');
// ---- minimal CRC32 ----
var CRC_TABLE = (function(){var t=new Array(256);for(var n=0;n<256;n++){var c=n;for(var k=0;k<8;k++)c=c&1?(0xEDB88320^(c>>>1)):(c>>>1);t[n]=c>>>0;}return t;})();
function crc32(buf){var c=0xFFFFFFFF;for(var i=0;i<buf.length;i++)c=CRC_TABLE[(c^buf[i])&0xFF]^(c>>>8);return (c^0xFFFFFFFF)>>>0;}
function rd2(b,o){return b[o]|(b[o+1]<<8);} function rd4(b,o){return (b[o]|(b[o+1]<<8)|(b[o+2]<<16)|(b[o+3]<<24))>>>0;}
function strBytes(s){var u=new Uint8Array(s.length);for(var i=0;i<s.length;i++)u[i]=s.charCodeAt(i)&0xFF;return u;}
// parse stored zip into entries [{name, data}]
function parseZip(u8){
  // find EOCD
  var i=u8.length-22;
  while(i>=0 && rd4(u8,i)!==0x06054b50) i--;
  var cdOff=rd4(u8,i+16), cnt=rd2(u8,i+10);
  var entries=[]; var p=cdOff;
  for(var e=0;e<cnt;e++){
    // central dir header
    var nameLen=rd2(u8,p+28), extraLen=rd2(u8,p+30), commLen=rd2(u8,p+32);
    var lho=rd4(u8,p+42);
    var name=''; for(var j=0;j<nameLen;j++) name+=String.fromCharCode(u8[p+46+j]);
    // read local header
    var lnameLen=rd2(u8,lho+26), lextraLen=rd2(u8,lho+28);
    var compSize=rd4(u8,p+20);
    var dataStart=lho+30+lnameLen+lextraLen;
    var data=u8.slice(dataStart, dataStart+compSize);
    entries.push({name:name, data:data});
    p+=46+nameLen+extraLen+commLen;
  }
  return entries;
}
function buildZip(entries){
  var locals=[], central=[], offset=0;
  function num2(n){return [n&0xFF,(n>>8)&0xFF];} function num4(n){return [n&0xFF,(n>>8)&0xFF,(n>>16)&0xFF,(n>>>24)&0xFF];}
  for(var k=0;k<entries.length;k++){
    var en=entries[k]; var nameB=strBytes(en.name); var crc=crc32(en.data); var sz=en.data.length;
    var lh=[].concat([0x50,0x4b,0x03,0x04],num2(20),num2(0),num2(0),num2(0),num2(0),num4(crc),num4(sz),num4(sz),num2(nameB.length),num2(0));
    locals.push(new Uint8Array(lh)); locals.push(nameB); locals.push(en.data);
    var cd=[].concat([0x50,0x4b,0x01,0x02],num2(20),num2(20),num2(0),num2(0),num2(0),num2(0),num4(crc),num4(sz),num4(sz),num2(nameB.length),num2(0),num2(0),num2(0),num2(0),num4(0),num4(offset));
    central.push(new Uint8Array(cd)); central.push(nameB);
    offset+=lh.length+nameB.length+sz;
    en._crc=crc; en._sz=sz;
  }
  var cdStart=offset; var cdLen=0;
  for(var c=0;c<central.length;c++) cdLen+=central[c].length;
  var eocd=new Uint8Array([].concat([0x50,0x4b,0x05,0x06],num2(0),num2(0),num2(entries.length),num2(entries.length),num4(cdLen),num4(cdStart),num2(0)));
  var parts=locals.concat(central,[eocd]); var total=0; for(var t=0;t<parts.length;t++) total+=parts[t].length;
  var out=new Uint8Array(total); var off=0; for(var u=0;u<parts.length;u++){out.set(parts[u],off);off+=parts[u].length;}
  return out;
}
function injectFreeze(u8){
  var entries=parseZip(u8);
  var dec=new TextDecoder('utf-8'), enc=new TextEncoder();
  var paneXml='<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/>';
  for(var i=0;i<entries.length;i++){
    if(/^xl\/worksheets\/sheet\d+\.xml$/.test(entries[i].name)){
      var xml=dec.decode(entries[i].data);
      xml=xml.replace(/<sheetView( [^>]*?)\/>/, '<sheetView$1>'+paneXml+'</sheetView>')
             .replace(/(<sheetView( [^>]*?)>)(?!<pane)/, function(m,p1){ return p1.indexOf(paneXml)>=0?p1: p1; });
      entries[i].data=enc.encode(xml);
    }
  }
  return buildZip(entries);
}
// build a workbook
var ws=X.utils.aoa_to_sheet([['Субъект','rdfs:label'],[':A','x'],[':B','y']]);
ws['!autofilter']={ref:'A1:B1'};
var wb=X.utils.book_new(); X.utils.book_append_sheet(wb,ws,'Процесс');
var u8=X.write(wb,{type:'array',bookType:'xlsx',compression:false});
var mod=injectFreeze(new Uint8Array(u8));
require('fs').writeFileSync('frozen.xlsx',Buffer.from(mod));
// re-read with SheetJS to validate zip integrity
var wb2=X.read(mod,{type:'array'});
console.log('reopened sheets:', wb2.SheetNames);
var dec=new TextDecoder('utf-8');
var entries=parseZip(mod).filter(e=>/sheet1\.xml$/.test(e.name));
var xml=dec.decode(entries[0].data);
console.log('pane present:', xml.includes('<pane'), '| autofilter present:', xml.includes('autoFilter'));
