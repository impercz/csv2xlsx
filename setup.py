from os.path import dirname, join
from setuptools import setup

VERSION = '0.1'

install_requires = []

with open(join(dirname(__file__), 'requirements.pip')) as req_file:
    for l in req_file.readlines():
        l = l.strip()
        if l and not l.startswith('#'):
            install_requires.append(l)

with open(join(dirname(__file__), 'README.rst')) as readme_file:
    long_description = readme_file.read()

setup_requires = ['setuptools']

setup(
    name='csv2xslx',
    version=VERSION,
    description='',
    long_description=long_description,
    author='Vlada Macek',
    author_email='macek@sandbox.cz',
    license='BSD',
    url='https://github.com/impercz/csv2xslx',
    py_modules=['csv2xlsx'],
    entry_points={
        'console_scripts': [
            'csv2xlsx = csv2xlsx:main',
        ],
    },
    zip_safe=False,
    classifiers=[
        "Development Status :: 4 - Beta",
        "Environment :: Console",
        "Intended Audience :: Developers",
        "License :: OSI Approved :: BSD License",
        "Operating System :: OS Independent",
        "Programming Language :: Python",
        "Programming Language :: Python :: 2",
        "Programming Language :: Python :: 2.7",
        "Topic :: Utilities",
    ],

    install_requires=install_requires,
    setup_requires=setup_requires
)
