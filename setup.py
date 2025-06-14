from setuptools import setup, find_packages

setup(
    name='access_tools',
    version='0.2',
    packages=find_packages(),
    install_requires=[
        'pandas',
        'pyodbc',
        'pywin32'
    ],
)
