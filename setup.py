import sys
from setuptools import find_packages, setup


is_64bit_python = sys.maxsize == 2**63 - 1
packages = ["pywechat"]
entry_points = {}
if is_64bit_python:
    packages = find_packages(include=["pywechat", "pywechat.*", "pyweixin", "pyweixin.*", "pyweixin_gui", "pyweixin_gui.*"])
    entry_points = {
        'console_scripts': [
            'autowechat=pyweixin_gui.app:main',
        ]
    }

setup(name='pywechat127',
version='1.9.9',
author='Hello-Mr-Crab',
author_email='3083256475@qq.com',
classifiers=[
"Operating System :: Microsoft :: Windows",
],
platforms=["Windows"],
description=f'A Powerful Windows-PC-Wechat automation Tool',
long_description=open('README.md','r',encoding='utf-8').read(),
long_description_content_type='text/markdown',  
url='https://github.com/Hello-Mr-Crab/pywechat',
#64位操作系统都安装,32位操作系统只安装pywechat
packages=packages,
license='LGPL',
keywords=['rpa','windows','wechat','automation'],
install_requires=[
'emoji>=2.14.1','PyAutoGUI>=0.9.54','pycaw>=20240210','pywin32>=308','pywin32-ctypes>=0.2.2','pywinauto>=0.6.8','psutil>=5.9.5','pillow>=10.4.0']
,
extras_require={
    'gui': ['PySide6>=6.6.0', 'openpyxl>=3.1.0'],
    'dev-gui': ['PySide6>=6.6.0', 'openpyxl>=3.1.0', 'PyInstaller>=6.0.0'],
},
entry_points=entry_points
)
'''
Author:Hello-Mr-Crab
Contributor:Viper,Chanpoe,mrhan1993,nmhjklnm,clen1,guanjt3
'''
