# setup.py
from distutils.core import setup
import py2app
data_files = ['GRE.xls', 'TOEFL.xls']
options_py2app = {
        'py2app' : {
           	'includes': ['xlrd']
            }
        }
setup(app=["words_denote.py"],options = options_py2app,data_files=data_files,setup_requires=['py2app'])

