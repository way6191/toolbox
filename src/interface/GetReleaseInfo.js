const fs = require('fs');
const path = require('path');
const sd = require('silly-datetime');

function getAll(filePath) {
  let out = [];
  getReleaseInfo(filePath, out);
  out.sort((item1, item2) => {
    var filePath1 = item1.filePath.toUpperCase(); // 大文字と小文字を無視する
    var filePath2 = item2.filePath.toUpperCase(); // 大文字と小文字を無視する
    if (filePath1 < filePath2) {
      return -1;
    }
    if (filePath1 > filePath2) {
      return 1;
    }
  });
  return out;
}

function getReleaseInfo(filePath, out) {
  let files = fs.readdirSync(filePath);

  files.forEach(filename => {
    //ファイルパース取得
    var filedir = path.join(filePath, filename);
    //ファイルstat取得
    var stats = fs.statSync(filedir);

    var isFile = stats.isFile(); //ファイル
    var isDir = stats.isDirectory(); //フォルダー
    if (isFile) {
      let info = {};
      info.filePath = filePath;
      info.fileName = filename;
      info.fileTime = sd.format(stats.mtime, 'YYYY/MM/DD HH:mm');
      info.fileSize = stats.size;

      out.push(info);
    }

    if (isDir) {
      getReleaseInfo(filedir, out);
    }
  });

}

module.exports = getAll;