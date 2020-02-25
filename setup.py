from setuptools import setup, find_packages

with open("README.md", "r") as fh:
    long_description = fh.read()

setup(
    name='laundry',
    version='2020.2.2 ',
    author='Tom Ackland',
    author_email='ackland.thomas@gmail.com',
    url='https://github.com/tomackl/laundry',
    description='Folding spreadsheets into neat shapes.',
    long_description=long_description,
    long_description_content_type="text/markdown",
    python_requires=">=3.6",
    packages=find_packages(where='src'),
    package_dir={'': 'src'},
    include_package_data=True,
    install_requires=[
        'Click',
        'Pandas',
        'python-docx',
        'pyjanitor',
        'colorama',
    ],
    entry_points={
        'console_scripts': [
            'laundry = laundry.laundry_cli:cli',
        ],
    },
    classifiers=[
        'Development Status :: 4 - Beta',
        'Operating System :: MacOS :: MacOS X',
        'Operating System :: Microsoft :: Windows :: Windows 10',
        'License :: OSI Approved :: Apache Software License',
        'Programming Language :: Python',
        'Programming Language :: Python :: 3.6',
        'Programming Language :: Python :: 3.7',
    ],
)

