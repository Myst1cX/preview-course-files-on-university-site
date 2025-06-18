// ==UserScript==
// @name         Preview .DOCX and PDF Files on VIS
// @namespace    http://tampermonkey.net/
// @version      3.1
// @description  Preview DOCX files using blob + mammoth.js viewer with theme switcher. Other docs open via Google Docs Viewer embed. PDF is natively supported by the browser.
// @author       Myst1cX
// @match        *://*/*
// @grant        GM_xmlhttpRequest
// @connect      *
// @homepageURL  https://github.com/Myst1cX/preview-course-files-on-university-site
// @supportURL   https://github.com/Myst1cX/preview-course-files-on-university-site/issues
// @updateURL    https://raw.githubusercontent.com/Myst1cX/preview-course-files-on-university-site/main/vis-file-preview.user.js
// @downloadURL  https://raw.githubusercontent.com/Myst1cX/preview-course-files-on-university-site/main/vis-file-preview.user.js
// ==/UserScript==

(function() {
    'use strict';

    const supportedExtensions = [
        '.doc', '.docx',
        '.ppt', '.pptx',
        '.xls', '.xlsx',
        '.rtf', '.txt',
        '.odt', '.ods', '.odp'
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

    async function openDocxBlobViewer(url) {
        try {
            const blob = await fetchFileAsBlob(url, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
            const arrayBuffer = await blob.arrayBuffer();

            const mammothJsUrl = 'https://unpkg.com/mammoth/mammoth.browser.min.js';

            // Convert arrayBuffer to a JSON-serializable array for inline use
            const uint8Array = new Uint8Array(arrayBuffer);
            const uint8ArrayJson = JSON.stringify(Array.from(uint8Array));

            const html = `
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8" />
<title>DOCX Preview</title>
<style id="theme-style">
    /* Dropdown always fixed */
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

        // Recreate ArrayBuffer from JSON array
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

    const links = document.querySelectorAll('a[href]');

    links.forEach(link => {
        if (link.nextSibling?.classList?.contains('preview-icon-btn')) return;

        const rawHref = link.getAttribute('href');
        if (!rawHref) return;

        const fullUrl = resolveUrl(rawHref);
        if (!fullUrl || !hasSupportedExtension(fullUrl)) return;

        const ext = fullUrl.split('?')[0].toLowerCase();
        const isDocx = ext.endsWith('.docx');

        const btn = document.createElement('button');
        btn.className = 'preview-icon-btn';
        btn.type = 'button';
        btn.title = isDocx ? 'Preview DOCX Document' : 'Preview Document (Google Docs Viewer)';
        btn.innerHTML = eyeIconSVG;

        btn.addEventListener('click', e => {
            e.preventDefault();
            e.stopPropagation();

            if (isDocx) {
                openDocxBlobViewer(fullUrl);
            } else {
                const previewWindow = window.open('', '_blank');
                if (!previewWindow) {
                    alert('Popup blocked! Please allow popups for this site.');
                    return;
                }
                const encoded = encodeURIComponent(fullUrl);
                previewWindow.location.href = `https://docs.google.com/gview?embedded=true&url=${encoded}`;
            }
        });

        link.parentNode.insertBefore(btn, link.nextSibling);
    });
})();
