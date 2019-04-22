from setuptools import setup, find_packages

setup(
    name='templar',
    version='2019.01.00',
    packages=find_packages(where='src'),
    package_dir={'': 'src'},
    # py_modules=['templar'],
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
)

