var todo_path = 'C:\\Users\\username\\Dropbox\\skypebot-todo';
var todo_file = 'todo.txt';
var todo_regex = /^(?:todo|とど|トド)[: 　。]/i;
var post_times = ['09:00', '13:00', '17:00'];

// framework
String.prototype.repeat = function(length){
  var ret = [];
  for(var i = 0; i < length; i++){ ret.push(this); }
  return ret.join('');
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
    var lines = text.split(/\r|\n|\r\n/);
    var output = [];
    for(var i = 0; i < lines.length; i++){
      var line = lines[i];
      if(todo_regex.test(line)){
        output.push(line);
      }
    }
    if(output.length != 0){
      chat.SendMessage(output.join('\n'));
    }
  }
}

function make_fullpath(chat){
  var fso = new ActiveXObject('Scripting.FileSystemObject');
  var name = chat.FriendlyName.replace(/[\\\/\:\*\?\"<>\|]/g, '_');
  return fso.BuildPath(todo_path, name + '_' + todo_file);
}
