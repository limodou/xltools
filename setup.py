__doc__ = """
=====================
Xltools
=====================

:Author: Limodou <limodou@gmail.com>

About Xltools
----------------

Xltools is used to read and write excel via excel template.

"""

from setuptools import setup, find_packages

setup(name='xltools',
    version='0.1',
    description="Extrace doc information from P8 project",
    long_description=__doc__,
    classifiers=[
        "Development Status :: 2 - Pre-Alpha",
        "Intended Audience :: Developers",
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python",
        "Operating System :: OS Independent",
        "Programming Language :: Python :: 2.7",
    ],
    packages= find_packages(),
    platforms = 'any',
    keywords='excel template',
    author='limodou',
    author_email='limodou@gmail.com',
    url='',
    license='MIT',
    include_package_data=True,
    zip_safe=False,
    entry_points = {
        'console_scripts': [
            'xltools = xltools:main',
        ],
    },
    install_requires=['openpyxl']
)
