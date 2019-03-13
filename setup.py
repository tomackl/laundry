from setuptools import setup, find_packages

setup(
    name='templar',
    version='0.1',
    packages=find_packages(),
    include_package_data=True,
    install_requires=[
        'Click',
    ],
    # TODO: complete the entry points below.
    # entry_points='''
    #     [console_scripts]
    #     yourscript=yourpackage.scripts.yourscript:cli
    # ''',
)
