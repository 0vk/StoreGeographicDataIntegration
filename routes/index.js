var express = require('express'),
  fs = require('fs'),
  formidable = require('formidable'),
  path = require('path'),
  _handle = require('./handle'),
  ldap = require('./util');
var router = express.Router();

/* login page : GET */
router.get('/login', function(req, res, next) {
  if (req.session.username) {
    res.redirect('/');
  } else {
    res.render('login', {
      title: 'Store Geographic Data Integration',
      msg: null
    });
  }
});

/* login page : POST */
router.post('/login', function(req, res, next) {
  var username, password;
  username = req.body.username;
  password = req.body.password;
  console.log(username, password);
  if ('' === username || '' === password) {
    res.render('login', {
      title: 'Store Geographic Data Integration',
      msg: 'The username or password is not correct, Authentication failed!'
    });
  }
  var p = ldap.authenticate(username, password);
  p.then(function(result) {
    req.session.username = username;
    res.redirect('/');
  }, function(error) {
    res.render('login', {
      title: 'Store Geographic Data Integration',
      msg: 'The username or password is not correct, Authentication failed!'
    });
  });
});

router.get('/logout', function(req, res) {
  req.session = null;
  res.redirect('/');
});

/* GET home page. */
router.get('/', function(req, res, next) {
  console.log(next);
  if (req.session.username) {
    var path = 'resources';
    if (fs.existsSync(path)) {
      var dirList = fs.readdirSync(path);
      dirList.forEach(function(fileName) {
        fs.unlinkSync(path + '/' + fileName);
      });
    }
    res.render('index', {
      title: 'Store Geographic Data Integration'
    });
  } else {
    res.redirect('login');
  }
});

/* POST upload. */
router.post('/upload', function(req, res, next) {
  // parse a file upload
  var form = new formidable.IncomingForm(),
    files = [],
    fields = [],
    docs = [];

  //存放目录
  // form.uploadDir = 'resources/';

  form.on('field', function(field, value) {
    //console.log(field, value);
    fields.push([field, value]);
  }).on('file', function(field, file) {
    if (/MDM_Store.*?/.test(file.name)) {
      fs.renameSync(file.path, 'resources/wb1.xlsx');
    }
    if (/CHINA OPEN.*?/.test(file.name)) {
      fs.renameSync(file.path, 'resources/wb2.xlsx');
    }
    if (/OPENCOCHINA.*?/.test(file.name)) {
      fs.renameSync(file.path, 'resources/wb3.xlsx');
    }
  }).on('end', function() {
    res.writeHead(200, {
      'content-type': 'text/plain'
    });
    var out = {
      Resopnse: {
        'result-code': 0,
        timeStamp: new Date(),
      },
      files: docs
    };
    var sout = JSON.stringify(out);
    res.end(sout);
  });

  form.parse(req, function(err, fields, files) {
    err && console.log('formidabel error : ' + err);

    console.log('parsing done');
  });
});

/* POST download. */
router.get('/download/:fileName', function(req, res, next) {
  var fileName = req.params.fileName;
  var filePath = path.join(__dirname, fileName);
  _handle.excel2dta();
  _handle.export2file();
  console.log('download');
  var file = 'resources/output.xlsx';
  res.download(file);
  // res.render('index', { title: 'Store geographic data integration' });
});
module.exports = router;
