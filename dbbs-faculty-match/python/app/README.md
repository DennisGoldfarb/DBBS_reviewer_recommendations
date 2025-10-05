# Embedded Python application

Place the Python modules and packages that should ship with the desktop
application in this directory. The build pipeline copies everything under
`python/app/` into the CPython embeddable runtime's `app/` folder so it can be
invoked via the bundled `python.exe` sidecar.
