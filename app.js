var express = require('express')
var app = express()

app.use(express.static('public'))

/*
app.get('/', function(req, res) {
  res.send('Hello World!')
})
*/

//app.use('/', require('./index'))

app.listen(3000, function () {
  console.log('Example app listening on port 3000!')
})
