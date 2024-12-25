/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */
import MarkdownIt from "markdown-it";


Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        console.log("Office is ready in Word");
        document.getElementById("convert-button").addEventListener("click", convertMarkdownToWord);
    }
});

async function convertMarkdownToWord() {
    const markdownInput = document.getElementById("markdown-input").value;
    const md = new MarkdownIt();
    const htmlContent = md.render(markdownInput);
    // console.log(htmlContent);
    try {
        await Word.run(async (context) => {
            const docBody = context.document.body;
            docBody.clear();
            docBody.insertHtml(htmlContent, Word.InsertLocation.start);
            await context.sync();
            console.log("Markdown converted to Word.");
        });
    } catch (error) {
        console.error("Error converting Markdown to Word:", error);
    }
}
