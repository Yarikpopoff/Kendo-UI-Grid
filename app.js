var express = require('express');
var app = express();
var server = require('http').Server(app);
var favicon = require('serve-favicon');

app.set('port', process.env.PORT || 3000);

server.listen(app.get('port'), function() {
	console.log("Listening on " + app.get('port'));
});

app.use(express.static(__dirname));

app.use(favicon(__dirname + '/favicon.ico'));

/*app.get('/', function (req, res) {
  res.sendfile(__dirname + '/index.html');
});

app.get('/Resume-En-Junior-JS-Developer-PopovYN.pdf', function (req, res) {
  res.sendfile(__dirname + '/Resume-En-Junior-JS-Developer-PopovYN.pdf');
});*/