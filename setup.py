#-*- coding:utf-8 -*-

#############################################
# File Name: setup.py
# Author: Jimmy Wong
# Mail: twong@126.com
# Created Time:  2018-12-11 19:17:34
#############################################

from setuptools import setup, find_packages

setup(
    name = "pdexcel",
    version = "0.0.1",
    keywords = ("pip", "SICA","featureextraction"),
    description = "Export DataFrame into an excel file",
    long_description = "An easy way to export DataFrame objects as tables and charts to excel file.",
    license = "MIT Licence",

    url = "https://github.com/UncleJiong/pdexcel",
    author = "Jimmy Wong",
    author_email = "twong@126.com",

    packages = find_packages(),
    include_package_data = True,
    platforms = "any",
    install_requires = ["pandas", "xlsxwriter", "matplotlib"]
)