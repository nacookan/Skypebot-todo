var todo_path = 'C:\\Users\\username\\Dropbox\\skypebot-todo';
var todo_file = 'todo.txt';
var todo_regex = /^([0-9ÇO-ÇX]*)([ Å@]?)(todo|Ç∆Ç«|ÉgÉh)([: Å@ÅB]+)(.+?)([ Å@])/i;
var countdown_time = '00:00';
var post_times = ['09:00', '13:00', '17:00'];

// framework
String.prototype.repeat = function(length){
  var ret = [];
  for(var i = 0; i < length; i++){ ret.push(this); }
  return ret.join('');
}
String.prototype.to_narrow_num = function(){
  return this.replace(/ÇO/g, '0').replace(/ÇP/g, '1').replace(/ÇQ/g, '2').replace(/ÇR/g, '3').replace(/ÇS/g, '4')
             .replace(/ÇT/g, '5').replace(/ÇU/g, '6').replace(/ÇV/g, '7').replace(/ÇW/g, '8').replace(/ÇX/g, '9');
}
Number.prototype.pad0 = function(length){
  return ('0'.repeat(length) + this.toString()).slice(-1 * length);
}
Date.prototype.hhmm = function(){
  return this.getHours().pad0(2) + ':' + this.getMinutes().pad0(2);
}
Array.prototype.indexOf = function(o){
  for(var i in this){ if(this[i] == o) return i; }
  return -1;
}

// skype4com
var skype = new ActiveXObject('Skype4COM.Skype');
WScript.ConnectObject(skype, 'Skype_');
skype.Attach();
function Skype_MessageStatus(msg, status){
  if(status == skype.Convert.TextToChatMessageStatus('RECEIVED')){
    receive(msg);
  }
}

// main infinite loop
var last_hhmm = null;
while(true){
  WScript.Sleep(1000);
  var hhmm = (new Date()).hhmm();
  if(hhmm == last_hhmm) continue;

  last_hhmm = hhmm;
  if(0 <= post_times.indexOf(hhmm)){
    post(skype);
  }
  if(countdown_time == hhmm){
    countdown();
  }
}

function receive(msg){
  var body = msg.Body;
  if(todo_regex.test(body)){
    var chat = msg.Chat;
    var fso = new ActiveXObject('Scripting.FileSystemObject');
    var fullpath = make_fullpath(chat);
    var stream = fso.OpenTextFile(fullpath, 8, true, -2);
    stream.WriteLine(body);
    stream.WriteLine('');
    stream.Close();
    chat.SendMessage('OK.');
    log('received.');
  }
}

function post(skype){
  var chats = new Enumerator(skype.Chats);
  for(; !chats.atEnd(); chats.moveNext()){
    var chat = chats.item();
    var fso = new ActiveXObject('Scripting.FileSystemObject');
    var fullpath = make_fullpath(chat);
    if(!fso.FileExists(fullpath)) continue;
    var stream = fso.OpenTextFile(fullpath, 1, false, -2);
    var text = '';
    if(!stream.AtEndOfStream){
      text = stream.ReadAll();
    }
    stream.Close();
    var lines = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
    var map = {};
    var keys = [];
    for(var i = 0; i < lines.length; i++){
      var line = lines[i];
      if(todo_regex.test(line)){
        if(RegExp.$1 == ''){
          if(!map[RegExp.$5]){
            map[RegExp.$5] = [];
            keys.push(RegExp.$5);
          }
          map[RegExp.$5].push(line.substr(0, 30));
        }
      }
    }
    if(keys.length != 0){
      for(var i = 0; i < keys.length; i++){
        chat.SendMessage(map[keys[i]].join('\n'));
        log('posted.');
      }
    }
  }
}

function countdown(){
  var chats = new Enumerator(skype.Chats);
  for(; !chats.atEnd(); chats.moveNext()){
    var chat = chats.item();
    var fso = new ActiveXObject('Scripting.FileSystemObject');
    var fullpath = make_fullpath(chat);
    if(!fso.FileExists(fullpath)) continue;
    var stream = fso.OpenTextFile(fullpath, 1, false, -2);
    var text = '';
    if(!stream.AtEndOfStream){
      text = stream.ReadAll();
    }
    stream.Close();
    var lines = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
    var output = [];
    for(var i = 0; i < lines.length; i++){
      var line = lines[i];
      if(todo_regex.test(line)){
        line = line.replace(todo_regex, function(string, g1, g2, g3, g4, g5, g6){
          if(g1 != ''){
            num = parseInt(g1.to_narrow_num(), 10);
            num--;
            if(0 < num){
              g1 = num.toString();
            }else{
              g1 = '';
            }
          }
          return g1 + g2 + g3 + g4 + g5 + g6;
        });
      }
      output.push(line);
    }
    var stream = fso.CreateTextFile(fullpath, true, false);
    stream.WriteLine(output.join('\r\n'));
    stream.Close();
    log('countdown.');
  }
}

function make_fullpath(chat){
  var fso = new ActiveXObject('Scripting.FileSystemObject');
  var name = chat.FriendlyName.replace(/[\\\/\:\*\?\"<>\|]/g, '_');
  return fso.BuildPath(todo_path, name + '_' + todo_file);
}

function log(message){
  var d = new Date();
  var line =
    d.getFullYear().toString() + '-' +
    (d.getMonth() + 1).toString() + '-' +
    d.getDate().toString() + ' ' +
    d.getHours().toString() + ':' +
    d.getMinutes().toString() + ':' +
    d.getSeconds().toString() + ' ' +
    message;
  WScript.echo(line);
}
