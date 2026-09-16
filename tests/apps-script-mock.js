/* A stand-in for the Apps Script services, faithful enough to catch the
   mistakes that matter: merges, sheet naming, property races, trigger counts. */
var CALLS = {};
function count(name){ CALLS[name] = (CALLS[name]||0)+1; }

function Range(sheet, row, col, nr, nc){
  this.s = sheet; this.r = row; this.c = col; this.nr = nr; this.nc = nc;
}
var STYLE = ['setVerticalAlignment','setFontFamily','setFontSize','setWrap','setFontWeight',
  'setBackground','setFontColor','setFontStyle','setBorder','setNumberFormat','setHorizontalAlignment'];
STYLE.forEach(function(m){ Range.prototype[m] = function(){ count('style'); return this; }; });
Range.prototype.getA1Notation = function(){
  return 'R' + this.r + 'C' + this.c + ':R' + (this.r+this.nr-1) + 'C' + (this.c+this.nc-1);
};
Range.prototype.overlapsPartialMerge_ = function(){
  var self = this;
  return this.s.merges.some(function(m){
    var inter = !(m.r+m.nr-1 < self.r || m.r > self.r+self.nr-1 ||
                  m.c+m.nc-1 < self.c || m.c > self.c+self.nc-1);
    if (!inter) return false;
    var contained = m.r >= self.r && m.c >= self.c &&
                    m.r+m.nr-1 <= self.r+self.nr-1 && m.c+m.nc-1 <= self.c+self.nc-1;
    return !contained;
  });
};
Range.prototype.getValues = function(){
  var out = [];
  for (var i=0;i<this.nr;i++){
    var row = [];
    for (var j=0;j<this.nc;j++) row.push(this.s.cell(this.r+i, this.c+j));
    out.push(row);
  }
  return out;
};
Range.prototype.getValue = function(){ return this.s.cell(this.r, this.c); };
Range.prototype.setValues = function(v){
  count('setValues');
  if (this.overlapsPartialMerge_()) throw new Error('Cannot write to a partially merged range');
  if (v.length !== this.nr) throw new Error('setValues: wrong row count '+v.length+' vs '+this.nr);
  for (var i=0;i<this.nr;i++){
    if (v[i].length !== this.nc) throw new Error('setValues: wrong column count '+v[i].length+' vs '+this.nc);
    for (var j=0;j<this.nc;j++) this.s.put(this.r+i, this.c+j, v[i][j]);
  }
  return this;
};
Range.prototype.setValue = function(x){ this.s.put(this.r, this.c, x); return this; };
Range.prototype.clearContent = function(){
  for (var i=0;i<this.nr;i++) for (var j=0;j<this.nc;j++) this.s.put(this.r+i, this.c+j, '');
  return this;
};
Range.prototype.merge = function(){
  count('merge');
  this.s.merges.push({r:this.r,c:this.c,nr:this.nr,nc:this.nc}); return this;
};
Range.prototype.breakApart = function(){
  var self = this;
  this.s.merges = this.s.merges.filter(function(m){
    return (m.r+m.nr-1 < self.r || m.r > self.r+self.nr-1 ||
            m.c+m.nc-1 < self.c || m.c > self.c+self.nc-1);
  });
  return this;
};

function RangeList(ranges){ this.ranges = ranges; }
STYLE.forEach(function(m){ RangeList.prototype[m] = function(){ count('style'); return this; }; });

function Protection(desc){ this.desc = desc; this.removed = false; }
Protection.prototype.setDescription = function(d){ this.desc = d; return this; };
Protection.prototype.getDescription = function(){ return this.desc; };
Protection.prototype.setWarningOnly = function(){ return this; };
Protection.prototype.remove = function(){ count('protect.remove'); this.removed = true; };

function Sheet(name){
  this.name = name; this.grid = {}; this.merges = []; this.hidden = false;
  this.protections = []; this.maxCols = 26; this.widths = {};
}
Sheet.prototype.key_ = function(r,c){ return r+','+c; };
Sheet.prototype.cell = function(r,c){ var v = this.grid[this.key_(r,c)]; return v === undefined ? '' : v; };
Sheet.prototype.put = function(r,c,v){ this.grid[this.key_(r,c)] = v; if (c > this.maxCols) this.maxCols = c; };
Sheet.prototype.getName = function(){ return this.name; };
Sheet.prototype.setName = function(n){ count('setName'); this.name = n; return this; };
Sheet.prototype.getMaxColumns = function(){ return this.maxCols; };
Sheet.prototype.getLastRow = function(){
  var max = 0; for (var k in this.grid){ if (this.grid[k] === '') continue;
    var r = Number(k.split(',')[0]); if (r > max) max = r; } return max;
};
Sheet.prototype.getLastColumn = function(){
  var max = 0; for (var k in this.grid){ if (this.grid[k] === '') continue;
    var c = Number(k.split(',')[1]); if (c > max) max = c; } return max;
};
Sheet.prototype.getRange = function(r,c,nr,nc){
  return new Range(this, r, c, nr === undefined ? 1 : nr, nc === undefined ? 1 : nc);
};
Sheet.prototype.getRangeList = function(list){
  var self = this;
  return new RangeList(list.map(function(a1){
    var m = a1.match(/^R(\d+)C(\d+):R(\d+)C(\d+)$/);
    return new Range(self, +m[1], +m[2], +m[3]-+m[1]+1, +m[4]-+m[2]+1);
  }));
};
Sheet.prototype.getDataRange = function(){
  return new Range(this, 1, 1, Math.max(this.getLastRow(),1), Math.max(this.getLastColumn(),1));
};
Sheet.prototype.clear = function(){ count('clear'); this.grid = {}; return this; };
Sheet.prototype.clearFormats = function(){ return this; };
Sheet.prototype.appendRow = function(v){
  var r = this.getLastRow()+1;
  for (var i=0;i<v.length;i++) this.put(r, i+1, v[i]);
};
Sheet.prototype.isSheetHidden = function(){ return this.hidden; };
Sheet.prototype.showSheet = function(){ this.hidden = false; return this; };
Sheet.prototype.hideSheet = function(){ this.hidden = true; return this; };
Sheet.prototype.getProtections = function(){
  return this.protections.filter(function(p){ return !p.removed; });
};
Sheet.prototype.protect = function(){
  count('protect.create');
  var p = new Protection(''); this.protections.push(p); return p;
};
Sheet.prototype.setColumnWidth = function(c,w){ count('width'); this.widths[c]=w; return this; };
Sheet.prototype.setColumnWidths = function(c,n,w){ count('width'); return this; };
Sheet.prototype.setFrozenRows = function(){ return this; };
Sheet.prototype.setFrozenColumns = function(){ return this; };

function Spreadsheet(){ this.sheets = []; this.active = null; this.toasts = []; }
Spreadsheet.prototype.getSheets = function(){ return this.sheets.slice(); };
Spreadsheet.prototype.getSheetByName = function(n){
  for (var i=0;i<this.sheets.length;i++) if (this.sheets[i].name === n) return this.sheets[i];
  return null;
};
Spreadsheet.prototype.insertSheet = function(n){
  count('insertSheet');
  if (this.getSheetByName(n)) throw new Error('A sheet named "'+n+'" already exists');
  var sh = new Sheet(n); this.sheets.push(sh); return sh;
};
Spreadsheet.prototype.deleteSheet = function(sh){
  count('deleteSheet');
  this.sheets = this.sheets.filter(function(s){ return s !== sh; });
};
Spreadsheet.prototype.setActiveSheet = function(sh){
  if (!sh) throw new Error('setActiveSheet(null)');
  this.active = sh; return sh;
};
Spreadsheet.prototype.moveActiveSheet = function(pos){
  var sh = this.active;
  this.sheets = this.sheets.filter(function(s){ return s !== sh; });
  this.sheets.splice(pos-1, 0, sh);
};
Spreadsheet.prototype.getActiveSheet = function(){ return this.active || this.sheets[0]; };
Spreadsheet.prototype.getSpreadsheetTimeZone = function(){ return 'America/New_York'; };
Spreadsheet.prototype.toast = function(msg){ this.toasts.push(msg); };

var SS = new Spreadsheet();
var SpreadsheetApp = {
  getActiveSpreadsheet: function(){ return SS; },
  getActive: function(){ return SS; },
  flush: function(){ count('flush'); },
  ProtectionType: { SHEET: 'SHEET' },
  BorderStyle: { SOLID: 'SOLID' },
  getUi: function(){ throw new Error('no UI in this context'); }
};

var PROPS = {};
var PropertiesService = {
  getDocumentProperties: function(){
    return {
      getProperty: function(k){ count('prop.get'); return PROPS[k] === undefined ? null : PROPS[k]; },
      setProperty: function(k,v){ count('prop.set'); PROPS[k] = String(v); return this; },
      deleteProperty: function(k){ count('prop.del'); delete PROPS[k]; return this; },
      getProperties: function(){ count('prop.all'); var o={}; for (var k in PROPS) o[k]=PROPS[k]; return o; },
      setProperties: function(o){ count('prop.setMany'); for (var k in o) PROPS[k]=String(o[k]); return this; }
    };
  }
};

var LOCKS = {};
function mkLock(name){
  return {
    tryLock: function(){ if (LOCKS[name]) return false; LOCKS[name]=true; return true; },
    waitLock: function(){ if (LOCKS[name]) throw new Error('lock busy'); LOCKS[name]=true; },
    releaseLock: function(){ LOCKS[name]=false; }
  };
}
var LockService = { getScriptLock: function(){ return mkLock('script'); },
                    getDocumentLock: function(){ return mkLock('document'); } };

var TRIGGERS = [];
var ScriptApp = {
  getProjectTriggers: function(){ count('trigger.list'); return TRIGGERS.slice(); },
  deleteTrigger: function(t){ count('trigger.delete'); TRIGGERS = TRIGGERS.filter(function(x){ return x!==t; }); },
  newTrigger: function(fn){
    return { timeBased: function(){ return {
      everyMinutes: function(){ return this; },
      everyHours:   function(){ return this; },
      after:        function(){ return this; },
      create: function(){
        count('trigger.create');
        if (TRIGGERS.length >= 20) throw new Error('This script has too many triggers.');
        var t = { getHandlerFunction: function(){ return fn; } };
        TRIGGERS.push(t); return t;
      }
    }; } };
  }
};

var Utilities = {
  formatDate: function(d){ return d.toISOString().replace('T',' ').substring(0,16); }
};
var ContentService = {
  createTextOutput: function(s){ return { setMimeType: function(){ return { _body: s }; } }; },
  MimeType: { JSON: 'JSON' }
};
var DriveApp = { getRootFolder: function(){ throw new Error('no Drive in tests'); } };
var console = { error: function(m){ print('   [log] ' + m); }, log: function(m){ print('   [log] ' + m); } };
