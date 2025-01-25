from distutils.core import setup
import py2exe

setup(
    windows=[{'script': 'main_function.py'}],
    options={
        'py2exe': 
        {
            'includes': ['lxml.etree', 'lxml._elementpath', 'gzip', 'Tkinter', 'tkFont', 'tkFileDialog', 'webbrowser', 'xlrd'],
        }
    }
)
