# Third-party components

This application uses unmodified third-party packages:

- Python: PSF license, https://www.python.org/psf/license/
- Pillow: HPND license, https://github.com/python-pillow/Pillow
- Qt / PySide6 / Shiboken: LGPLv3 / GPL / commercial licensing as supplied by the Qt project, https://doc.qt.io/qtforpython-6/licenses.html
- PyInstaller: GPL with a bootloader exception permitting generated application distribution, https://pyinstaller.org/en/stable/license.html

Qt libraries are dynamically linked and unpacked at runtime. Full application source and the build scripts are supplied so that users can replace the Qt libraries and rebuild this executable. The application does not restrict replacing those libraries or debugging modifications to them. Qt/PySide source and corresponding release archives are available from https://download.qt.io/official_releases/ and https://code.qt.io/ .

The executable bundles license notices where supplied by the dependency packages. Source code for this application is supplied alongside the initial handoff. No third-party library source has been modified.
