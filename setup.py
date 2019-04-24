from setuptools import setup, find_packages

setup(
    name='templar',
    version='2019.01.00',
    author='Tom Ackland',
    description='A app for bashing .xlsx files into .docx.',
    python_requires=">=3.7",
    packages=find_packages(where='src'),  # since it hasn't been tested on anything else.
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
            # 'templar = templar.src.templar:cli',
        ],
    },
    classifiers=[
        # 'Development Status :: 5 - Production/Stable',
        'Operating System :: MacOS :: MacOS X',
        # 'Operating System :: Microsoft :: Windows',
        # 'License :: OSI Approved :: MIT License',
        'Programming Language :: Python',
        'Programming Language :: Python :: 3.7',
    ],
)

