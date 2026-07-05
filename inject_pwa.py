import sys
content = open('index.html', 'r', encoding='utf-8').read()
replacement = '<link rel="manifest" href="manifest.json">\n    <script>\n        if("serviceWorker" in navigator) {\n            window.addEventListener("load", () => {\n                navigator.serviceWorker.register("sw.js");\n            });\n        }\n    </script>\n</head>'
if 'rel="manifest"' not in content:
    content = content.replace('</head>', replacement)
    open('index.html', 'w', encoding='utf-8').write(content)
    print("PWA tags injected successfully.")
else:
    print("PWA tags already exist.")
