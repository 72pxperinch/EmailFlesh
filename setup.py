from setuptools import setup
import os
import sys

# Determine the correct framework paths for macOS
if sys.platform == 'darwin':
    possible_tk_locations = [
        'Library/Frameworks/Python.framework/Versions/3.11/lib/python3.11/tkinter',
    ]
    
    possible_tcl_locations = [
        '/System/Library/Frameworks/Tcl.framework',
        '/Library/Frameworks/Tcl.framework',
        '/usr/local/opt/tcl-tk/lib',  # Homebrew location
        os.path.expanduser('~/Library/Frameworks/Tcl.framework'),
    ]
    
    # Find the first existing Tk framework
    tk_framework = None
    for loc in possible_tk_locations:
        if os.path.exists(loc):
            tk_framework = loc
            break
            
    # Find the first existing Tcl framework
    tcl_framework = None
    for loc in possible_tcl_locations:
        if os.path.exists(loc):
            tcl_framework = loc
            break
            
    if not tk_framework or not tcl_framework:
        print("Warning: Could not find Tk/Tcl frameworks")
        tk_framework = tcl_framework = ''

APP = ['EmailFlesh.py']
DATA_FILES = []
OPTIONS = {
    'argv_emulation': False,
    'packages': [
        'tkinter',
        'json',
        'email',
        'imaplib',
        'threading',
        'logging',
    ],
    'includes': [
        'email.header',
        'tkinter.ttk',
        'traceback',
    ],
    'frameworks': [tk_framework, tcl_framework] if sys.platform == 'darwin' and tk_framework and tcl_framework else [],
    'iconfile': 'icon.png',
    'plist': {
        'CFBundleName': 'EmailFlesh',
        'CFBundleDisplayName': 'EmailFlesh',
        'CFBundleGetInfoString': "Download email attachments",
        'CFBundleIdentifier': "com.svntytoopixels.emailflesh",
        'CFBundleVersion': "1.0.0",
        'CFBundleShortVersionString': "1.0.0",
        'NSHumanReadableCopyright': u"Copyright © 2024, svntytoopixels, All Rights Reserved",
        'LSMinimumSystemVersion': '10.13',
        'NSHighResolutionCapable': True,
    }
}

setup(
    app=APP,
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
) 