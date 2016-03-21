# setup.py
from distutils.core import setup
import py2exe
data_files = ['GRE.xls', 'TOEFL.xls']
options_py2exe = {
        'py2exe' : {
            'includes': ['xlrd']
            }
        }
setup(console=["words_denote.py"],options = options_py2exe,data_files=data_files)
