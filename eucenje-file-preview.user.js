// ==UserScript==
// @name         Preview .DOCX and PDF Files on E-Ucenje
// @namespace    http://tampermonkey.net/
// @version      2.2
// @description  Preview PDFs and DOCX files using blob URLs
// @author       Myst1cX
// @match        *://e-ucenje.ff.uni-lj.si/*
// @grant        GM_xmlhttpRequest
// @connect      e-ucenje.ff.uni-lj.si
// @homepageURL  https://github.com/Myst1cX/preview-course-files-on-university-site
// @supportURL   https://github.com/Myst1cX/preview-course-files-on-university-site/issues
// @updateURL    https://raw.githubusercontent.com/Myst1cX/preview-course-files-on-university-site/main/eucenje-file-preview.user.js
// @downloadURL  https://raw.githubusercontent.com/Myst1cX/preview-course-files-on-university-site/main/eucenje-file-preview.user.js
// ==/UserScript==

(function () {
    'use strict';

    const isEUcenje = location.hostname.includes('e-ucenje.ff.uni-lj.si');
    const wordDocExts = ['.doc', '.docx'];
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

    function addPreviewButtons() {
        document.querySelectorAll('a[href]').forEach(link => {
            const href = link.getAttribute('href');
            const fullUrl = resolveUrl(href);
            if (!fullUrl) return;

            if (link.parentNode.classList.contains('preview-wrapper')) return;

            const ext = fullUrl.split('?')[0].toLowerCase();

            const isPreviewable = (
                isEUcenje && (ext.endsWith('.pdf') || wordDocExts.some(e => ext.endsWith(e)))
            ) || isMoodleResourceLink(fullUrl);

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
                if (isMoodleResourceLink(targetUrl)) {
                    targetUrl = await getFinalUrl(targetUrl);
                }

                if (targetUrl.toLowerCase().endsWith('.pdf')) {
                    await openPdfBlobViewer(targetUrl);
                } else if (targetUrl.toLowerCase().endsWith('.docx')) {
                    await openDocxWithMammoth(targetUrl);
                } else if (targetUrl.toLowerCase().endsWith('.doc')) {
                    alert('Preview not supported for .doc files. Please convert to .docx to view.');
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
