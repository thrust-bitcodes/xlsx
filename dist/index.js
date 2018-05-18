var create = require('./create.js');
var read = require('./read.js');
function writeBytesToFile(bytes, fileName) {
    var fos;

    try {
        fos = new java.io.FileOutputStream(fileName);
        fos.write(bytes);
    } finally {
        if (fos) {
            fos.close();
        }
    }
}

exports = {
    create: create,
    read: read,
    writeBytesToFile: writeBytesToFile
}