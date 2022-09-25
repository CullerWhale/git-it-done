const docx = require('docx@6.0.1');
const fs = require('fs');
const request = require('request@2.88.0');
const express = require("@runkit/runkit/express-endpoint/1.0.0");
const app = express(exports);
var testEl = document.querySelector('#testbtn');

const {
    Document,
    HorizontalPositionAlign,
    HorizontalPositionRelativeFrom,
    ImageRun,
    Media,
    Packer,
    Paragraph,
    VerticalPositionAlign,
    VerticalPositionRelativeFrom,
} = docx;

// https://stackoverflow.com/questions/12740659/downloading-images-with-node-js
const download = (uri, filename, callback) => {
  request.head(uri, (err, res, body) => {
    request(uri).pipe(fs.createWriteStream(filename)).on('close', callback);
  });
};

const URL = 'https://raw.githubusercontent.com/dolanmiu/docx/ccd655ef8be3828f2c4b1feb3517a905f98409d9/demo/images/cat.jpg';

app.get("/", (req, res) => {
    download(URL, 'cat.jpg', async () => {
        const doc = new Document({
            sections: [{
                children: [
                    new Paragraph("Hello World"),
                    new Paragraph({
                        children: [
                            new ImageRun({
                                data: fs.readFileSync("./cat.jpg"),
                                transformation: {
                                    width: 100,
                                    height: 100,
                                }
                            }),
                        ],
                    }),
                    new Paragraph({
                        children: [
                            new ImageRun({
                                data: fs.readFileSync("./cat.jpg"),
                                transformation: {
                                    width: 200,
                                    height: 200,
                                },
                                floating: {
                                    horizontalPosition: {
                                        offset: 1014400,
                                    },
                                    verticalPosition: {
                                        offset: 1014400,
                                    },
                                },
                            }),
                        ],
                    }),
                    new Paragraph({
                        children: [
                            new ImageRun({
                                data: fs.readFileSync("./cat.jpg"),
                                transformation: {
                                    width: 200,
                                    height: 200,
                                },
                                floating: {
                                    horizontalPosition: {
                                        relative: HorizontalPositionRelativeFrom.PAGE,
                                        align: HorizontalPositionAlign.RIGHT,
                                    },
                                    verticalPosition: {
                                        relative: VerticalPositionRelativeFrom.PAGE,
                                        align: VerticalPositionAlign.BOTTOM,
                                    },
                                },
                            }),
                        ],
                    }),
                ],
            }],
        });
        
        const b64string = await Packer.toBase64String(doc);

        res.setHeader('Content-Disposition', 'attachment; filename=My Document.docx');
        res.send(Buffer.from(b64string, 'base64'));
    });
});

var testfunction = function() {
    
    alert('test worked');

};


testEl.addEventListener('click', testfunction);