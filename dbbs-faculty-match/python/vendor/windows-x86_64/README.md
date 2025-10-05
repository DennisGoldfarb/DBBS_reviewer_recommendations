# Windows vendored packages

Place the Windows (x86_64) Python site-packages tree in this directory. You can
build it on a Windows development machine by installing the dependencies into a
virtual environment and copying its `Lib/site-packages` folder here. Make sure
all wheels are compatible with Python 3.11 and the 64-bit embeddable
interpreter.

If you download the official CPython embeddable distribution manually, the ZIP
file can also be saved in this directory as `python-3.11.9-embed-amd64.zip` so
that `npm run prepare-python` can reuse it without fetching it again.
