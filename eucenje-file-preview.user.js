// ==UserScript==
// @name         Preview PDFs, DOCX and PPTX files on E-Ucenje
// @namespace    http://tampermonkey.net/
// @version      2.3
// @description  The files are previewed in a new tab using blob URLs.
// @author       Myst1cX
// @match        *://e-ucenje.ff.uni-lj.si/*
// @grant        GM_xmlhttpRequest
// @connect      e-ucenje.ff.uni-lj.si
// @homepageURL  https://github.com/Myst1cX/uni-preview-course-files
// @supportURL   https://github.com/Myst1cX/uni-preview-course-files/issues
// @updateURL    https://raw.githubusercontent.com/Myst1cX/uni-preview-course-files/main/eucenje-file-preview.user.js
// @downloadURL  https://raw.githubusercontent.com/Myst1cX/uni-preview-course-files/main/eucenje-file-preview.user.js
// ==/UserScript==

(function () {
    'use strict';

    const isEUcenje = location.hostname.includes('e-ucenje.ff.uni-lj.si');
    const supportedExtensions = ['.doc', '.docx', '.pptx'];
    const pdfjsViewerBase = 'https://mozilla.github.io/pdf.js/web/viewer.html?file=';

    const isMoodleResourceLink = url => {
      try { return isEUcenje && new URL(url).pathname.includes('/mod/resource/view.php'); }
      catch { return false; }
    };


    function resolveUrl(url) {
        try { return new URL(url, location.href).href; }
        catch { return null; }
    }

    function getFinalUrl(url) {
        return new Promise(resolve => {
            GM_xmlhttpRequest({
                method: 'HEAD',
                url,
                onload: resp => {
                    const final = resp.finalUrl || resp.responseURL || url;
                    resolve(final);
                },
                onerror: () => resolve(url)
            });
        });
    }

    function fetchFileAsBlob(url, mime) {
        return new Promise((resolve, reject) => {
            GM_xmlhttpRequest({
                method: 'GET',
                url,
                responseType: 'arraybuffer',
                headers: { 'Accept': mime },
                onload: response => {
                    if (response.status >= 200 && response.status < 300) {
                        const blob = new Blob([response.response], { type: mime });
                        resolve(blob);
                    } else {
                        reject(new Error('Failed to fetch file, status: ' + response.status));
                    }
                },
                onerror: () => reject(new Error('Network error while fetching file')),
            });
        });
    }

    async function openPdfBlobViewer(url) {
        try {
            const blob = await fetchFileAsBlob(url, 'application/pdf');
            const blobUrl = URL.createObjectURL(blob);
            window.open(blobUrl, '_blank');
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
        const htmlContent = `<!DOCTYPE html>
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
    <h1>${fileName}</h1>
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



    async function openDocxWithMammoth(url) {
        try {
            const blob = await fetchFileAsBlob(url, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
            const arrayBuffer = await blob.arrayBuffer();

            const mammothJsUrl = 'https://unpkg.com/mammoth/mammoth.browser.min.js';

            const htmlContent = `
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>DOCX Preview</title>
    <style id="theme-style">
        /* Always keep dropdown fixed */
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

        /* Base styles applied with theme */
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
                body {
                    background: #fff;
                    color: #000;
                }
            \`,
            dark: \`
                body {
                    background: #121212;
                    color: #e0e0e0;
                }
            \`,
            sepia: \`
                body {
                    background: #f4ecd8;
                    color: #5b4636;
                }
            \`
        };

        const styleTag = document.getElementById('theme-style');
        const themeSwitcher = document.querySelector('.theme-switcher');

        function applyTheme(theme) {
            styleTag.innerHTML = \`
                /* Keep dropdown fixed */
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

                /* Base styles */
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

                /* Theme colors */
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

        const arrayBuffer = new Uint8Array(${JSON.stringify([...new Uint8Array(arrayBuffer)])}).buffer;
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

            const htmlBlob = new Blob([htmlContent], { type: 'text/html' });
            const viewerUrl = URL.createObjectURL(htmlBlob);
            window.open(viewerUrl, '_blank');

        } catch (e) {
            alert('Error loading DOCX preview:\n' + e.message);
        }
    }

    const eyeIcon = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"
        stroke-linecap="round" stroke-linejoin="round" width="24" height="24">
        <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"></path>
        <circle cx="12" cy="12" r="3"></circle></svg>`;

function getFileExtension(url) {
    const match = url.split('?')[0].match(/\.([a-z0-9]+)$/i);
    return match ? match[1].toLowerCase() : null;
}

function isForumAttachmentUrl(url) {
    return url.includes('pluginfile.php') && url.includes('/mod_forum/attachment/');
}

function addPreviewButtons() {
    document.querySelectorAll('a[href]').forEach(link => {
        const href = link.getAttribute('href');
        const fullUrl = resolveUrl(href);
        if (!fullUrl) return;

        if (link.parentNode.classList.contains('preview-wrapper')) return;

        const lowerUrl = fullUrl.toLowerCase();
        const ext = getFileExtension(lowerUrl);

const isPreviewable = (
    isEUcenje && (
        ext && supportedExtensions.includes(`.${ext}`) ||
        isMoodleResourceLink(fullUrl) ||
        isForumAttachmentUrl(fullUrl)
    )
);

        if (!isPreviewable) return;

        const wrapper = document.createElement('div');
        wrapper.className = 'preview-wrapper';

        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'preview-icon-btn';
        btn.title = 'Preview Document';
        btn.innerHTML = eyeIcon;

        btn.addEventListener('click', async e => {
            e.preventDefault();
            e.stopPropagation();

            let targetUrl = fullUrl;

            if (isMoodleResourceLink(targetUrl) || isForumAttachmentUrl(targetUrl)) {
                targetUrl = await getFinalUrl(targetUrl);
            }

            const finalExt = getFileExtension(targetUrl);

            if (finalExt === 'pdf') {
    await openPdfBlobViewer(targetUrl);
} else if (finalExt === 'docx' || finalExt === 'doc') {
    if (finalExt === 'doc') {
        alert('Preview not supported for .doc files. Please convert to .docx to view.');
    } else {
        await openDocxWithMammoth(targetUrl);
    }
} else if (finalExt === 'pptx' || finalExt === 'ppt') {
    if (finalExt === 'ppt') {
        alert('Preview not supported for .ppt files. Please convert to .pptx to view.');
    } else {
        await openPptxWithPptxJs(targetUrl);
    }
} else {
    alert('Preview not supported for this file type.');
}

        });

        link.parentNode.insertBefore(wrapper, link);
        wrapper.appendChild(btn);
        wrapper.appendChild(link);
    });
}


    if (!document.getElementById('preview-pdf-css')) {
        document.head.insertAdjacentHTML('beforeend', `
            <style id="preview-pdf-css">
                .preview-wrapper {
                    display: inline-flex;
                    flex-direction: row;
                    align-items: center;
                    margin-bottom: 16px;
                    position: relative;
                }
                .preview-icon-btn {
                    border: none;
                    background: none;
                    cursor: pointer;
                    color: #4285F4;
                    padding: 6px;
                    margin-right: 6px;
                    width: 30px;
                    height: 30px;
                    display: flex;
                    justify-content: center;
                    align-items: center;
                    z-index: 10;
                }
                .preview-icon-btn:hover {
                    color: #3367D6;
                }
                .preview-wrapper a {
                    position: relative;
                    z-index: 1;
                }
            </style>`);
    }

    function run() {
        addPreviewButtons();
        setTimeout(run, 2000);
    }

    run();
})();
