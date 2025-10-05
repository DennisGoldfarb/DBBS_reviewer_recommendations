# Vendored Python packages

This directory stores the site-packages trees that are bundled with the
application. Populate the subdirectory that matches the platform you are
packaging (for example `windows-x86_64/site-packages`). The contents are copied
verbatim into the embeddable Python runtime during `npm run prepare-python`.
