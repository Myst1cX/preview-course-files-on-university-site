// ==UserScript==
// @name         Preview PDFs, .DOCX, .PPTX and .XLSX files on VIS
// @namespace    http://tampermonkey.net/
// @version      4.0
// @description  The files are previewed in a new tab using blob URLs.
// @author       Myst1cX
// @match        *://visff.uni-lj.si/*
// @grant        GM_xmlhttpRequest
// @connect      visff.uni-lj.si
// @homepageURL  https://github.com/Myst1cX/uni-preview-course-files
// @supportURL   https://github.com/Myst1cX/uni-preview-course-files/issues
// @updateURL    https://raw.githubusercontent.com/Myst1cX/uni-preview-course-files/main/vis-file-preview.user.js
// @downloadURL  https://raw.githubusercontent.com/Myst1cX/uni-preview-course-files/main/vis-file-preview.user.js
// ==/UserScript==

(function() {
    'use strict';

     const supportedExtensions = [
        '.docx',
        '.pdf',
        '.xlsx',
        '.pptx'
    ];

    const eyeIconSVG = `
        <svg xmlns="http://www.w3.org/2000/svg"
             viewBox="0 0 24 24" fill="none" stroke="currentColor"
             stroke-width="2" stroke-linecap="round" stroke-linejoin="round"
             width="16" height="16" style="vertical-align: middle;">
            <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"></path>
            <circle cx="12" cy="12" r="3"></circle>
        </svg>
    `;

    const style = document.createElement('style');
    style.textContent = `
        .preview-icon-btn {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            margin-left: 4px;
            padding: 0;
            border: none;
            background: none;
            cursor: pointer;
            color: #4285F4;
            vertical-align: middle;
            text-decoration: none;
        }
        .preview-icon-btn:hover {
            color: #3367D6;
        }
    `;
    document.head.appendChild(style);

    function resolveUrl(url) {
        try {
            return new URL(url, location.href).href;
        } catch {
            return null;
        }
    }

    function hasSupportedExtension(url) {
        try {
            const pathname = new URL(url).pathname.toLowerCase();
            return supportedExtensions.some(ext => pathname.endsWith(ext));
        } catch {
            return false;
        }
    }

    function fetchFileAsBlob(url, mime) {
        return new Promise((resolve, reject) => {
            GM_xmlhttpRequest({
                method: 'GET',
                url,
                responseType: 'arraybuffer',
                headers: { 'Accept': mime },
                onload: (res) => {
                    if (res.status >= 200 && res.status < 300) {
                        resolve(new Blob([res.response], { type: mime }));
                    } else {
                        reject(new Error(`Failed to fetch file: ${res.status}`));
                    }
                },
                onerror: () => reject(new Error('Network error')),
            });
        });
    }

async function openPdfBlobViewer(url) {
    try {
        const blob = await fetchFileAsBlob(url, 'application/pdf');
        const blobUrl = URL.createObjectURL(blob);

        const rawFileName = url.split('/').pop().split('?')[0];
        const fileName = decodeURIComponent(rawFileName);

        const viewerHtml = `
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8" />
    <title>${fileName}</title>
    <style>
        html, body {
            margin: 0;
            padding: 0;
            height: 100%;
            overflow: hidden;
        }
        iframe {
            border: none;
            width: 100%;
            height: 100%;
        }
    </style>
</head>
<body>
    <iframe src="${blobUrl}" allow="fullscreen"></iframe>
</body>
</html>`;

        const viewerBlob = new Blob([viewerHtml], { type: 'text/html' });
        const viewerBlobUrl = URL.createObjectURL(viewerBlob);

        window.open(viewerBlobUrl, '_blank');
    } catch (e) {
        alert('Error loading PDF:\n' + e.message);
    }
}

    async function openPptxWithPptxJs(url) {
    try {
        const blob = await fetchFileAsBlob(
            url,
            'application/vnd.openxmlformats-officedocument.presentationml.presentation'
        );
        const pptxBlobUrl = URL.createObjectURL(blob);

        const rawFileName = url.split('/').pop().split('?')[0];
        const fileName = decodeURIComponent(rawFileName);
        const htmlContent = `
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8" />
    <title>${fileName}</title>
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/gh/meshesha/PPTXjs/css/pptxjs.css" />
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/gh/meshesha/PPTXjs/css/nv.d3.min.css" />
    <style>
        body {
            font-family: sans-serif;
            margin: 0;
            background: #222;
            color: #eee;
            display: flex;
            flex-direction: column;
            align-items: center;
            line-height: 1.5;
        }
        #viewer {
            width: 100%;
            max-width: 90vw;
            height: 80vh;
            margin: 20px auto;
            overflow: auto;
            border: 1px solid #444;
            background-color: #111;
            position: relative;
            white-space: normal !important;
            word-wrap: break-word !important;
            overflow-wrap: break-word !important;
            hyphens: auto !important;
            padding: 16px 20px;
            box-sizing: border-box;
        }
        #viewer span.text-block {
            display: inline-block;
            max-width: 100%;
            white-space: normal !important;
            word-break: break-word !important;
            overflow-wrap: break-word !important;
            hyphens: auto !important;
            margin-right: 6px;
        }
    </style>
    <script src="https://cdn.jsdelivr.net/gh/meshesha/PPTXjs/js/jquery-1.11.3.min.js"></script>
    <script src="https://cdn.jsdelivr.net/gh/meshesha/PPTXjs/js/jquery.fullscreen-min.js"></script>
    <script src="https://cdn.jsdelivr.net/gh/meshesha/PPTXjs/js/jszip.min.js"></script>
    <script src="https://cdn.jsdelivr.net/gh/meshesha/PPTXjs/js/filereader.js"></script>
    <script src="https://cdn.jsdelivr.net/gh/meshesha/PPTXjs/js/d3.min.js"></script>
    <script src="https://cdn.jsdelivr.net/gh/meshesha/PPTXjs/js/nv.d3.min.js"></script>
    <script src="https://cdn.jsdelivr.net/gh/meshesha/PPTXjs/js/dingbat.js"></script>
    <script src="https://cdn.jsdelivr.net/gh/meshesha/PPTXjs/js/pptxjs.js"></script>
    <script src="https://cdn.jsdelivr.net/gh/meshesha/PPTXjs/js/divs2slides.js"></script>
</head>
<body>
    <div id="viewer"></div>
    <script>
        $("#viewer").pptxToHtml({
            pptxFileUrl: "${pptxBlobUrl}",
            slideMode: false,
            keyBoardShortCut: false,
            mediaProcess: true,
            themeProcess: true,
            slideType: "divs2slidesjs"
        });

        function fixTextNodes(node) {
            if (node.nodeType === Node.TEXT_NODE) {
                node.textContent = node.textContent.replace(/\\u00A0/g, ' ');
            } else if (node.nodeType === Node.ELEMENT_NODE) {
                node.querySelectorAll('span.text-block').forEach(el => {
                    el.style.whiteSpace = 'normal';
                    el.style.wordBreak = 'break-word';
                    el.style.overflowWrap = 'break-word';
                    el.style.hyphens = 'auto';
                    el.style.maxWidth = '100%';
                    el.style.display = 'inline-block';
                    el.style.marginRight = '6px';
                });
                node.childNodes.forEach(child => fixTextNodes(child));
            }
        }

        function fixSpanTypos() {
            document.querySelectorAll('span.text-block').forEach(el => {
                const span = document.createElement('span');
                span.className = el.className;
                span.style.cssText = el.style.cssText;
                el.childNodes.forEach(child => {
                    span.appendChild(child.cloneNode(true));
                });
                fixTextNodes(span);
                el.parentNode.replaceChild(span, el);
            });
        }

        function fixMissingSpacesBetweenSpans(container) {
            const spans = container.querySelectorAll('span.text-block');
            for (let i = 0; i < spans.length - 1; i++) {
                const current = spans[i];
                const next = spans[i + 1];
                const currentLastChar = current.textContent.slice(-1);
                const nextFirstChar = next.textContent.charAt(0);
                if (currentLastChar !== ' ' && nextFirstChar !== ' ') {
                    current.parentNode.insertBefore(document.createTextNode(' '), next);
                }
            }
        }

        let fixScheduled = false;
        function scheduleFixes() {
            if (fixScheduled) return;
            fixScheduled = true;
            requestAnimationFrame(() => {
                fixSpanTypos();
                fixTextNodes(document.getElementById('viewer'));
                fixMissingSpacesBetweenSpans(document.getElementById('viewer'));
                fixScheduled = false;
            });
        }

        const viewer = document.getElementById('viewer');
        const observer = new MutationObserver(mutations => {
            scheduleFixes();
        });

        observer.observe(viewer, {
            childList: true,
            subtree: true,
            characterData: true,
        });

        setTimeout(() => {
            fixSpanTypos();
            fixTextNodes(viewer);
            fixMissingSpacesBetweenSpans(viewer);
        }, 2000);
    </script>
</body>
</html>`;

        const htmlBlob = new Blob([htmlContent], { type: 'text/html' });
        const blobUrl = URL.createObjectURL(htmlBlob);
        window.open(blobUrl, '_blank');
    } catch (e) {
        alert('Error loading PPTX preview:\n' + e.message);
    }
}

    async function openDocxBlobViewer(url) {
        try {
            const blob = await fetchFileAsBlob(url, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
            const arrayBuffer = await blob.arrayBuffer();

            const mammothJsUrl = 'https://unpkg.com/mammoth/mammoth.browser.min.js';

            const rawFileName = url.split('/').pop().split('?')[0];
            const fileName = decodeURIComponent(rawFileName);

            const uint8Array = new Uint8Array(arrayBuffer);
            const uint8ArrayJson = JSON.stringify(Array.from(uint8Array));

            const html = `
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8" />
<title>${fileName}</title>
<style id="theme-style">
    .theme-switcher {
        position: fixed;
        top: 10px;
        right: 20px;
        z-index: 9999;
        font-size: 14px;
        padding: 4px 6px;
        border-radius: 4px;
        border: 1px solid #ccc;
        background: white;
        color: black;
        cursor: pointer;
    }
    body {
        font-family: sans-serif;
        padding: 20px;
        max-width: 800px;
        margin: auto;
        line-height: 1.6;
        word-wrap: break-word;
    }
    img {
        max-width: 100%;
        height: auto;
        display: block;
        margin: 1em 0;
    }
    h1, h2, h3 {
        margin-top: 1.5em;
    }
    table {
        border-collapse: collapse;
        width: 100%;
        overflow-x: auto;
    }
    table, th, td {
        border: 1px solid #ccc;
        padding: 6px;
    }
</style>
<script src="${mammothJsUrl}"></script>
</head>
<body>
    <select class="theme-switcher" aria-label="Select theme">
        <option value="light">Light</option>
        <option value="dark">Dark</option>
        <option value="sepia">Sepia</option>
    </select>
    <div id="output">Loading...</div>
    <script>
        const themes = {
            light: \`
                body { background: #fff; color: #000; }
            \`,
            dark: \`
                body { background: #121212; color: #e0e0e0; }
            \`,
            sepia: \`
                body { background: #f4ecd8; color: #5b4636; }
            \`
        };

        const styleTag = document.getElementById('theme-style');
        const themeSwitcher = document.querySelector('.theme-switcher');

        function applyTheme(theme) {
            styleTag.innerHTML = \`
                .theme-switcher {
                    position: fixed;
                    top: 10px;
                    right: 20px;
                    z-index: 9999;
                    font-size: 14px;
                    padding: 4px 6px;
                    border-radius: 4px;
                    border: 1px solid #ccc;
                    background: white;
                    color: black;
                    cursor: pointer;
                }
                body {
                    font-family: sans-serif;
                    padding: 20px;
                    max-width: 800px;
                    margin: auto;
                    line-height: 1.6;
                    word-wrap: break-word;
                }
                img {
                    max-width: 100%;
                    height: auto;
                    display: block;
                    margin: 1em 0;
                }
                h1, h2, h3 {
                    margin-top: 1.5em;
                }
                table {
                    border-collapse: collapse;
                    width: 100%;
                    overflow-x: auto;
                }
                table, th, td {
                    border: 1px solid #ccc;
                    padding: 6px;
                }
                \${themes[theme]}
            \`;
        }

        const savedTheme = localStorage.getItem('docxPreviewTheme') || 'light';
        themeSwitcher.value = savedTheme;
        applyTheme(savedTheme);

        themeSwitcher.addEventListener('change', e => {
            const selected = e.target.value;
            applyTheme(selected);
            localStorage.setItem('docxPreviewTheme', selected);
        });

        const uint8ArrayData = ${uint8ArrayJson};
        const arrayBuffer = new Uint8Array(uint8ArrayData).buffer;

        mammoth.convertToHtml({ arrayBuffer: arrayBuffer })
            .then(result => {
                document.getElementById('output').innerHTML = result.value;
            })
            .catch(err => {
                document.getElementById('output').innerText = 'Error rendering DOCX: ' + err.message;
            });
    </script>
</body>
</html>
`;

            const htmlBlob = new Blob([html], { type: 'text/html' });
            const blobUrl = URL.createObjectURL(htmlBlob);
            window.open(blobUrl, '_blank');
        } catch (e) {
            alert('Error loading DOCX preview:\n' + e.message);
        }
    }

    async function openXlsxWithSheetJs(url) {
    try {
        const blob = await fetchFileAsBlob(url, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        const arrayBuffer = await blob.arrayBuffer();

        const sheetJsUrl = 'https://unpkg.com/xlsx/dist/xlsx.full.min.js';

        const rawFileName = url.split('/').pop().split('?')[0];
        const fileName = decodeURIComponent(rawFileName);

        const htmlContent = `
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>${fileName}</title>
    <style id="theme-style">
        .theme-switcher {
            position: fixed;
            top: 10px;
            right: 20px;
            z-index: 9999;
            font-size: 14px;
            padding: 4px 6px;
            border-radius: 4px;
            border: 1px solid #ccc;
            background: white;
            color: black;
            cursor: pointer;
        }

        body {
            font-family: sans-serif;
            padding: 20px;
            max-width: 95vw;
            margin: auto;
            line-height: 1.6;
            overflow-x: auto;
        }

        table {
            border-collapse: collapse;
            width: 100%;
            margin-top: 1em;
        }
        th, td {
            border: 1px solid #ccc;
            padding: 6px;
            text-align: left;
        }
    </style>
    <script src="${sheetJsUrl}"></script>
</head>
<body>
    <select class="theme-switcher" aria-label="Select theme">
        <option value="light">Light</option>
        <option value="dark">Dark</option>
        <option value="sepia">Sepia</option>
    </select>
    <div id="output">Loading...</div>
    <script>
        const themes = {
            light: \`
                body { background: #fff; color: #000; }
            \`,
            dark: \`
                body { background: #121212; color: #e0e0e0; }
                table, th, td { border-color: #444; }
            \`,
            sepia: \`
                body { background: #f4ecd8; color: #5b4636; }
            \`
        };

        const styleTag = document.getElementById('theme-style');
        const themeSwitcher = document.querySelector('.theme-switcher');

        function applyTheme(theme) {
            styleTag.innerHTML = \`
                .theme-switcher {
                    position: fixed;
                    top: 10px;
                    right: 20px;
                    z-index: 9999;
                    font-size: 14px;
                    padding: 4px 6px;
                    border-radius: 4px;
                    border: 1px solid #ccc;
                    background: white;
                    color: black;
                    cursor: pointer;
                }

                body {
                    font-family: sans-serif;
                    padding: 20px;
                    max-width: 95vw;
                    margin: auto;
                    line-height: 1.6;
                    overflow-x: auto;
                }

                table {
                    border-collapse: collapse;
                    width: 100%;
                    margin-top: 1em;
                }
                th, td {
                    border: 1px solid #ccc;
                    padding: 6px;
                    text-align: left;
                }

                \${themes[theme]}
            \`;
        }

        const savedTheme = localStorage.getItem('xlsxPreviewTheme') || 'light';
        themeSwitcher.value = savedTheme;
        applyTheme(savedTheme);

        themeSwitcher.addEventListener('change', e => {
            const selected = e.target.value;
            applyTheme(selected);
            localStorage.setItem('xlsxPreviewTheme', selected);
        });

        const arrayBuffer = new Uint8Array(${JSON.stringify([...new Uint8Array(arrayBuffer)])}).buffer;
        const data = new Uint8Array(arrayBuffer);
        const workbook = XLSX.read(data, { type: "array" });

        const output = document.getElementById('output');
        output.innerHTML = "";

        workbook.SheetNames.forEach(sheetName => {
            const html = XLSX.utils.sheet_to_html(workbook.Sheets[sheetName]);
            const sheetDiv = document.createElement("div");
            sheetDiv.innerHTML = "<h2>" + sheetName + "</h2>" + html;
            output.appendChild(sheetDiv);
        });
    </script>
</body>
</html>
`;

        const htmlBlob = new Blob([htmlContent], { type: 'text/html' });
        const viewerUrl = URL.createObjectURL(htmlBlob);
        window.open(viewerUrl, '_blank');

    } catch (e) {
        alert('Error loading XLSX preview:\n' + e.message);
    }
}

    const links = document.querySelectorAll('a[href]');

    links.forEach(link => {
        if (link.nextSibling?.classList?.contains('preview-icon-btn')) return;

        const rawHref = link.getAttribute('href');
        if (!rawHref) return;

        const fullUrl = resolveUrl(rawHref);
        if (!fullUrl || !hasSupportedExtension(fullUrl)) return;

        const hrefLower = fullUrl.toLowerCase();
        const ext = fullUrl.split('?')[0].toLowerCase();
        const isDocx = ext.endsWith('.docx');
        const isPdf = ext.endsWith('.pdf');
        const isXlsx = hrefLower.endsWith('.xlsx');
        const isPptx = hrefLower.endsWith('.pptx');

        if (!isDocx && !isPdf && !isXlsx && !isPptx) return;

        const btn = document.createElement('button');
        btn.className = 'preview-icon-btn';
        btn.type = 'button';

        if (isDocx) {
            btn.title = 'Preview DOCX Document';
            btn.innerHTML = eyeIconSVG;
            btn.addEventListener('click', e => {
                e.preventDefault();
                e.stopPropagation();
                openDocxBlobViewer(fullUrl);
            });
        } else if (isPdf) {
            btn.title = 'Preview PDF Document';
            btn.innerHTML = eyeIconSVG;
            btn.addEventListener('click', e => {
                e.preventDefault();
                e.stopPropagation();
                openPdfBlobViewer(fullUrl);
            });
        } else if (isXlsx) {
            btn.title = 'Preview XLSX Spreadsheet';
            btn.innerHTML = eyeIconSVG;
            btn.addEventListener('click', e => {
                e.preventDefault();
                e.stopPropagation();
                openXlsxBlobViewer(fullUrl);
            });
        } else if (isPptx) {
            btn.title = 'Preview PPTX Presentation';
            btn.innerHTML = eyeIconSVG;
            btn.addEventListener('click', e => {
                e.preventDefault();
                e.stopPropagation();
                openPptxBlobViewer(fullUrl);
            });
        }

        link.parentNode.insertBefore(btn, link.nextSibling);
    });
})();
