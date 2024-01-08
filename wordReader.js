const mammoth = require('mammoth');
const fs = require('fs');

mammoth.convertToHtml({ path: './origin.docx' }).then(result => {
    const html = result.value;
    const messages = result.messages;
    console.log(html, messages);
});