# -*- coding: utf-8 -*-
from setuptools import setup, find_packages


with open("README.md", "r") as fh:
    long_description = fh.read()

setup(
    name='aioxlsxstream',
    version='1.0.1',
    description='Asynchronous xlsx files generator',
    long_description=long_description,
    long_description_content_type="text/markdown",
    author='lososkin',
    author_email='yalososka@gmail.com',
    url='https://github.com/lososkin/aioxlsxstream',
    packages=find_packages(exclude=['tests']),
    keywords='async xlsx streaming',
    install_requires="asynczipstream>=1.0.1",
    classifiers=[
        "Programming Language :: Python",
        "License :: OSI Approved :: GNU General Public License v3 (GPLv3)",
        "Operating System :: OS Independent",
        "Topic :: System :: Archiving :: Compression",
    ],
)
