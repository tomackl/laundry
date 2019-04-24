from setuptools import setup, find_packages

with open("README.md", "r") as fh:
    long_description = fh.read()

setup(
    name='templar',
    version='2019.01.0',
    author='Tom Ackland',
    author_email='ackland dot thomas at gmail dot com',
    description='A app for bashing .xlsx files into .docx.',
    long_description=long_description,
    long_description_content_type="text/markdown",
    python_requires=">=3.7",
    packages=find_packages(where='src'),
    package_dir={'': 'src'},
    include_package_data=True,
    install_requires=[
        'Click',
        'Pandas',
        'python-docx',
        'pyjanitor',
    ],
    entry_points={
        'console_scripts': [
            'templar = templar:cli',
        ],
    },
    classifiers=[
        'Development Status :: 4 - Beta',
        'Operating System :: MacOS :: MacOS X',
        'License :: OSI Approved :: Apache Software License',
        'Programming Language :: Python',
        'Programming Language :: Python :: 3.7',
    ],
)

